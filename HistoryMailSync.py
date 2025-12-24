"""
历史邮件同步脚本
用于将收件箱和发件箱中已手动处理的邮件同步到数据库
避免系统重复发送
"""

import poplib
import email
from email.parser import Parser
from email.policy import default
import time
import os
import logging
from datetime import datetime, timedelta
import sqlite3
import re
import sys
from email.header import decode_header

# 导入现有模块的函数
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 配置参数（应该从主配置文件读取，这里先使用默认值）
email_address = "zhang.peiying@coscoshipping.com"
password = "29CqCF5ZRUbcqIKG"
pop3_server = "mail.coscoshipping.com"
pop3_port = 995

# 数据库文件（使用现有的数据库）
IMPORT_DB_FILE = 'processed_emails_import.db'
EXPORT_DB_FILE = 'processed_emails.db'

# 关键词配置（应该从配置文件读取）
IMPORT_KEYWORDS = ['Calcium Nitrate', 'Calcium Nitrate Tetrahydrate', 'Magnesium Nitrate Hexahydrate']
EXPORT_KEYWORDS = ['Calcium Nitrate', 'Calcium Nitrate Tetrahydrate', 'Magnesium Nitrate Hexahydrate']

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('history_sync.log', encoding='utf-8')
    ]
)

class HistoryMailSync:
    def __init__(self, config_path='config.ini'):
        """
        初始化历史邮件同步器
        
        Args:
            config_path: 配置文件路径
        """
        self.config_path = config_path
        self.email_config = self.load_email_config()
        
        # 统计信息
        self.stats = {
            'total_emails': 0,
            'import_synced': 0,
            'export_synced': 0,
            'already_processed': 0,
            'no_attachment': 0,
            'error': 0
        }
        
    def load_email_config(self):
        """加载邮箱配置"""
        try:
            import configparser
            config = configparser.ConfigParser()
            
            if os.path.exists(self.config_path):
                config.read(self.config_path, encoding='utf-8')
                
                email_config = {
                    'email_address': config.get('email', '进口邮箱地址', fallback=email_address),
                    'password': config.get('email', '进口邮箱密码', fallback=password),
                    'pop3_server': config.get('email', 'pop3服务器', fallback=pop3_server),
                    'pop3_port': config.getint('email', 'pop3端口', fallback=pop3_port),
                    'keywords': {
                        'import': [],
                        'export': []
                    }
                }
                
                # 从配置文件加载关键词
                if 'keywords' in config:
                    for key in config['keywords']:
                        value = config['keywords'][key]
                        if value.strip():
                            if key.startswith('进口关键词'):
                                email_config['keywords']['import'].append(value.strip())
                            elif key.startswith('出口关键词'):
                                email_config['keywords']['export'].append(value.strip())
                
                # 如果配置文件中没有关键词，使用默认值
                if not email_config['keywords']['import']:
                    email_config['keywords']['import'] = IMPORT_KEYWORDS
                if not email_config['keywords']['export']:
                    email_config['keywords']['export'] = EXPORT_KEYWORDS
                    
                return email_config
                
        except Exception as e:
            logging.error(f"加载配置文件失败: {e}")
            
        # 返回默认配置
        return {
            'email_address': email_address,
            'password': password,
            'pop3_server': pop3_server,
            'pop3_port': pop3_port,
            'keywords': {
                'import': IMPORT_KEYWORDS,
                'export': EXPORT_KEYWORDS
            }
        }
    
    def decode_email_header(self, header):
        """解码邮件头"""
        if not header:
            return ""
        
        try:
            if isinstance(header, bytes):
                header = header.decode('utf-8', errors='ignore')
            
            if '=?' in header and '?=' in header:
                decoded_parts = []
                for part, charset in decode_header(header):
                    if isinstance(part, bytes):
                        if charset:
                            try:
                                decoded_parts.append(part.decode(charset, errors='ignore'))
                            except:
                                decoded_parts.append(part.decode('utf-8', errors='ignore'))
                        else:
                            decoded_parts.append(part.decode('utf-8', errors='ignore'))
                    else:
                        decoded_parts.append(str(part))
                return ''.join(decoded_parts)
            else:
                return header
        except Exception as e:
            logging.warning(f"解码邮件头失败: {e}")
            return header
    
    def extract_email_address(self, email_string):
        """从邮件字符串中提取邮箱地址"""
        try:
            decoded_string = self.decode_email_header(email_string)
            
            match = re.search(r'<([^>]+)>', decoded_string)
            if match:
                return match.group(1).strip()
            
            if re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', decoded_string.strip()):
                return decoded_string.strip()
            
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', decoded_string)
            if email_match:
                return email_match.group(0).strip()
            
            return decoded_string.strip()
        except Exception as e:
            logging.error(f"提取邮箱地址失败: {e}")
            return email_string
    
    def is_export_manifest(self, txt_content):
        """判断是否为出口舱单"""
        try:
            if not txt_content:
                return False
            
            sample = txt_content[:500] if len(txt_content) > 500 else txt_content
            
            if "00NCLCONTAINER LIST" in sample:
                return True
            
            if "00:IFCSUM:" in sample:
                return False
            
            lines = txt_content.split('\n')
            colon_count = 0
            total_lines_checked = min(20, len(lines))
            
            for i in range(total_lines_checked):
                line = lines[i]
                if ':' in line and line.count(':') >= 5:
                    colon_count += 1
            
            if colon_count >= 3:
                return False
            
            return False  # 简化判断，主要依赖已有逻辑
        except Exception as e:
            logging.error(f"判断出口舱单失败: {e}")
            return False
    
    def is_import_manifest(self, txt_content):
        """判断是否为进口舱单"""
        try:
            if not txt_content:
                return False
            
            sample = txt_content[:500] if len(txt_content) > 500 else txt_content
            
            if "00:IFCSUM:" in sample:
                return True
            
            if "00NCLCONTAINER LIST" in sample:
                return False
            
            lines = txt_content.split('\n')
            import_pattern_count = 0
            
            for line in lines[:30]:
                if line.startswith(('00:', '10:', '11:', '12:', '13:', '16:', '17:', '18:', '41:', '44:', '47:', '51:')):
                    import_pattern_count += 1
            
            if import_pattern_count >= 5:
                return True
            
            return False
        except Exception as e:
            logging.error(f"判断进口舱单失败: {e}")
            return False
    
    def check_keywords_in_text(self, text, keyword_type='import'):
        """检查文本中是否包含关键词"""
        if not text:
            return []
        
        keywords = self.email_config['keywords'][keyword_type]
        found_keywords = []
        
        for keyword in keywords:
            if keyword.lower() in text.lower():
                found_keywords.append(keyword)
        
        return found_keywords
    
    def get_email_attachments(self, msg):
        """获取邮件附件"""
        attachments = []
        
        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition"))
            
            if "attachment" in content_disposition:
                filename = part.get_filename()
                if filename:
                    decoded_filename = self.decode_email_header(filename)
                    
                    # 读取附件内容
                    try:
                        file_content = part.get_payload(decode=True)
                        if isinstance(file_content, bytes):
                            try:
                                file_content = file_content.decode('utf-8', errors='ignore')
                            except:
                                file_content = file_content.decode('gbk', errors='ignore')
                        
                        attachments.append({
                            'filename': decoded_filename,
                            'content': file_content
                        })
                    except Exception as e:
                        logging.error(f"读取附件失败: {e}")
        
        return attachments
    
    def sync_email_to_database(self, email_uid, msg, folder='inbox'):
        """
        将邮件同步到数据库
        
        Args:
            email_uid: 邮件UID
            msg: 邮件对象
            folder: 邮件所在文件夹（inbox/sent）
        """
        try:
            # 获取邮件基本信息
            subject = self.decode_email_header(msg.get('subject', '无主题'))
            from_header = self.decode_email_header(msg.get('from', '未知发件人'))
            from_addr = self.extract_email_address(from_header)
            to_header = self.decode_email_header(msg.get('to', ''))
            date = self.decode_email_header(msg.get('date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
            
            logging.info(f"处理邮件: {subject}")
            
            # 获取附件
            attachments = self.get_email_attachments(msg)
            
            if not attachments:
                self.stats['no_attachment'] += 1
                return False
            
            # 检查是否已处理（检查数据库）
            for db_file, db_type in [(IMPORT_DB_FILE, 'import'), (EXPORT_DB_FILE, 'export')]:
                if os.path.exists(db_file):
                    conn = sqlite3.connect(db_file)
                    cursor = conn.cursor()
                    
                    # 检查keyword_emails表
                    cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE email_uid = ?', (email_uid,))
                    if cursor.fetchone()[0] > 0:
                        logging.info(f"邮件已存在数据库: {email_uid}")
                        self.stats['already_processed'] += 1
                        conn.close()
                        return False
                    
                    conn.close()
            
            # 处理每个附件
            processed = False
            for attachment in attachments:
                if attachment['filename'].lower().endswith('.txt'):
                    txt_content = attachment['content']
                    
                    # 判断舱单类型并同步到对应数据库
                    if self.is_import_manifest(txt_content):
                        processed = self.sync_to_import_db(email_uid, subject, from_addr, date, 
                                                         attachment['filename'], txt_content)
                        if processed:
                            self.stats['import_synced'] += 1
                            break
                    
                    elif self.is_export_manifest(txt_content):
                        processed = self.sync_to_export_db(email_uid, subject, from_addr, date,
                                                         attachment['filename'], txt_content)
                        if processed:
                            self.stats['export_synced'] += 1
                            break
            
            if not processed:
                logging.info(f"邮件未包含可识别的舱单附件: {subject}")
            
            return processed
            
        except Exception as e:
            logging.error(f"同步邮件失败: {e}")
            self.stats['error'] += 1
            return False
    
    def sync_to_import_db(self, email_uid, subject, sender, date, filename, content):
        """同步到进口数据库"""
        try:
            self.ensure_sync_column_exists(IMPORT_DB_FILE)
            # 检查是否包含关键词
            found_keywords = self.check_keywords_in_text(content, 'import')
            matched_keywords_str = ','.join(found_keywords) if found_keywords else '历史同步'
            
            # 连接进口数据库
            conn = sqlite3.connect(IMPORT_DB_FILE)
            cursor = conn.cursor()
            
            # 确保表存在
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS keyword_emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_uid TEXT NOT NULL,
                sender TEXT NOT NULL,
                sender_address TEXT,
                subject TEXT NOT NULL,
                received_date TEXT,
                processed_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                matched_keywords TEXT NOT NULL,
                excel_sent INTEGER DEFAULT 1,
                txt_attachment TEXT,
                container_count INTEGER DEFAULT 0,
                attachment_names TEXT,
                sync_source TEXT DEFAULT 'history_sync',
                UNIQUE(email_uid)
            )
            ''')
            
            # 插入记录
            cursor.execute('''
            INSERT OR IGNORE INTO keyword_emails 
            (email_uid, sender, sender_address, subject, received_date, matched_keywords, 
             txt_attachment, container_count, attachment_names, sync_source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (email_uid, sender, sender, subject, date, matched_keywords_str,
                  filename, 0, filename, 'history_sync'))
            
            conn.commit()
            conn.close()
            
            logging.info(f"✅ 已同步到进口数据库: {subject}")
            return True
            
        except Exception as e:
            logging.error(f"同步到进口数据库失败: {e}")
            return False

    ##新增
    def ensure_sync_column_exists(self, db_file):
        """确保数据库表中有sync_source列"""
        try:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # 检查表是否存在
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails'")
            if not cursor.fetchone():
                logging.error(f"表keyword_emails不存在于{db_file}")
                conn.close()
                return False
            
            # 检查列是否存在
            cursor.execute("PRAGMA table_info(keyword_emails)")
            columns = [info[1] for info in cursor.fetchall()]
            
            if 'sync_source' not in columns:
                logging.info(f"在{db_file}中添加sync_source列...")
                cursor.execute("ALTER TABLE keyword_emails ADD COLUMN sync_source TEXT DEFAULT ''")
                conn.commit()
                logging.info(f"✅ 已添加sync_source列到{db_file}")
            
            conn.close()
            return True
            
        except Exception as e:
            logging.error(f"确保sync_source列存在失败: {e}")
            return False

    
    def sync_to_export_db(self, email_uid, subject, sender, date, filename, content):
        """同步到出口数据库"""
        try:
            self.ensure_sync_column_exists(IMPORT_DB_FILE)
            # 检查是否包含关键词
            found_keywords = self.check_keywords_in_text(content, 'export')
            matched_keywords_str = ','.join(found_keywords) if found_keywords else '历史同步'
            
            # 连接出口数据库
            conn = sqlite3.connect(EXPORT_DB_FILE)
            cursor = conn.cursor()
            
            # 确保表存在
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS keyword_emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                email_uid TEXT NOT NULL,
                sender TEXT NOT NULL,
                sender_address TEXT,
                subject TEXT NOT NULL,
                received_date TEXT,
                processed_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                matched_keywords TEXT NOT NULL,
                excel_sent INTEGER DEFAULT 1,
                txt_attachment TEXT,
                container_count INTEGER DEFAULT 0,
                attachment_names TEXT,
                sync_source TEXT DEFAULT 'history_sync',
                UNIQUE(email_uid)
            )
            ''')
            
            # 插入记录
            cursor.execute('''
            INSERT OR IGNORE INTO keyword_emails 
            (email_uid, sender, sender_address, subject, received_date, matched_keywords, 
             txt_attachment, container_count, attachment_names, sync_source)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (email_uid, sender, sender, subject, date, matched_keywords_str,
                  filename, 0, filename, 'history_sync'))
            
            conn.commit()
            conn.close()
            
            logging.info(f"✅ 已同步到出口数据库: {subject}")
            return True
            
        except Exception as e:
            logging.error(f"同步到出口数据库失败: {e}")
            return False
    
    def sync_folder(self, folder='inbox', max_emails=100):
        """
        同步指定文件夹的邮件
        
        Args:
            folder: 邮件文件夹（inbox/sent）
            max_emails: 最大同步邮件数量
        """
        try:
            logging.info(f"开始同步{folder}文件夹...")
            
            # 连接邮箱服务器
            server = poplib.POP3_SSL(self.email_config['pop3_server'], 
                                     self.email_config['pop3_port'], timeout=30)
            
            # 登录
            server.user(self.email_config['email_address'])
            server.pass_(self.email_config['password'])
            
            # 获取邮件统计
            email_count, total_size = server.stat()
            logging.info(f"{folder}文件夹中共有 {email_count} 封邮件")
            
            # 限制处理数量
            process_count = min(email_count, max_emails)
            logging.info(f"将处理最近 {process_count} 封邮件")
            
            # 获取邮件UID列表
            response, uid_list, _ = server.uidl()
            uids = []
            for uid_line in uid_list:
                uid_str = uid_line.decode('utf-8')
                parts = uid_str.split()
                if len(parts) >= 2:
                    uids.append(parts[1])
            
            # 处理邮件（从最新开始）
            for i in range(1, process_count + 1):
                try:
                    uid = uids[i-1]
                    
                    # 获取邮件内容
                    response, lines, _ = server.retr(i)
                    msg_content = b'\r\n'.join(lines).decode('utf-8', errors='ignore')
                    msg = Parser(policy=default).parsestr(msg_content)
                    
                    # 同步到数据库
                    self.sync_email_to_database(uid, msg, folder)
                    
                    self.stats['total_emails'] += 1
                    
                    # 显示进度
                    if i % 10 == 0:
                        logging.info(f"已处理 {i}/{process_count} 封邮件")
                    
                    # 避免请求过快
                    time.sleep(0.1)
                    
                except Exception as e:
                    logging.error(f"处理第 {i} 封邮件失败: {e}")
                    self.stats['error'] += 1
            
            # 关闭连接
            server.quit()
            logging.info(f"{folder}文件夹同步完成")
            
            return True
            
        except Exception as e:
            logging.error(f"同步{folder}文件夹失败: {e}")
            return False
    
    # 在 HistoryMailSync 类的 sync_all_folders 方法中，确保进度回调被正确使用
    def sync_all_folders(self, max_emails=100, progress_callback=None):
        """同步所有文件夹的邮件，支持进度回调"""
        try:
            # 重置统计
            self.stats = {
                'total_emails': 0,
                'import_synced': 0,
                'export_synced': 0,
                'already_processed': 0,
                'no_attachment': 0,
                'error': 0
            }
            
            # 连接邮箱服务器
            server = poplib.POP3_SSL(self.email_config['pop3_server'], 
                                    self.email_config['pop3_port'])
            
            # 登录
            server.user(self.email_config['email_address'])
            server.pass_(self.email_config['password'])
            
            # 获取邮件统计
            email_count, total_size = server.stat()
            logging.info(f"收件箱中共有 {email_count} 封邮件")
            
            # 限制处理数量
            process_count = min(email_count, max_emails)
            logging.info(f"将处理最近 {process_count} 封邮件")
            
            # 更新进度：连接成功
            if progress_callback:
                if not progress_callback(0, process_count, "连接邮箱服务器成功"):
                    server.quit()
                    return {
                        'total': 0,
                        'import': 0,
                        'export': 0,
                        'already': 0,
                        'no_attachment': 0,
                        'error': 0,
                        'status': 'cancelled',
                        'message': '用户取消同步'
                    }
            
            # 获取邮件UID列表
            response, uid_list, _ = server.uidl()
            uids = []
            for uid_line in uid_list:
                uid_str = uid_line.decode('utf-8', errors='ignore')
                parts = uid_str.split()
                if len(parts) >= 2:
                    uids.append(parts[1])
            
            # 处理邮件（从最新开始）
            for i in range(1, process_count + 1):
                try:
                    # 检查是否需要停止
                    if progress_callback and not progress_callback(i-1, process_count, f"处理第 {i}/{process_count} 封邮件"):
                        logging.info("用户请求停止同步")
                        break
                    
                    # 获取邮件UID
                    uid = uids[i-1] if (i-1) < len(uids) else str(i)
                    
                    # 获取邮件内容
                    response, lines, _ = server.retr(i)
                    msg_content = b'\r\n'.join(lines).decode('utf-8', errors='ignore')
                    msg = Parser(policy=default).parsestr(msg_content)
                    
                    # 同步到数据库
                    self.sync_email_to_database(uid, msg, 'inbox')
                    
                    self.stats['total_emails'] += 1
                    
                    # 显示进度
                    if i % 10 == 0:
                        logging.info(f"已处理 {i}/{process_count} 封邮件")
                    
                    # 更新进度回调
                    if progress_callback:
                        progress_callback(i, process_count, f"已处理 {i}/{process_count} 封邮件")
                    
                    # 避免请求过快
                    time.sleep(0.1)
                    
                except Exception as e:
                    logging.error(f"处理第 {i} 封邮件失败: {e}")
                    self.stats['error'] += 1
            
            # 关闭连接
            server.quit()
            
            # 更新进度：完成
            if progress_callback:
                progress_callback(process_count, process_count, "同步完成")
            
            logging.info(f"同步完成，总计处理 {self.stats['total_emails']} 封邮件")
            
            return {
                'total': self.stats['total_emails'],
                'import': self.stats['import_synced'],
                'export': self.stats['export_synced'],
                'already': self.stats['already_processed'],
                'no_attachment': self.stats['no_attachment'],
                'error': self.stats['error'],
                'status': 'completed',
                'message': '同步完成'
            }
            
        except Exception as e:
            logging.error(f"同步文件夹失败: {e}")
            import traceback
            logging.error(traceback.format_exc())
            
            if progress_callback:
                progress_callback(0, 0, f"同步失败: {str(e)}")
            
            raise
    
    def get_sync_summary(self):
        """获取同步摘要信息"""
        return {
            'total': self.stats['total_emails'],
            'import': self.stats['import_synced'],
            'export': self.stats['export_synced'],
            'already': self.stats['already_processed'],
            'no_attachment': self.stats['no_attachment'],
            'error': self.stats['error']
        }


def main():
    """主函数"""
    import argparse
    
    parser = argparse.ArgumentParser(description='历史邮件同步工具')
    parser.add_argument('--config', '-c', default='config.ini', help='配置文件路径')
    parser.add_argument('--max-emails', '-m', type=int, default=100, help='每个文件夹最大同步邮件数')
    parser.add_argument('--inbox-only', action='store_true', help='仅同步收件箱')
    
    args = parser.parse_args()
    
    sync = HistoryMailSync(args.config)
    sync.sync_all_folders(args.max_emails)


if __name__ == "__main__":
    main()
    