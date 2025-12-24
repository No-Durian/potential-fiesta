import sqlite3
import os
import csv
import json
from datetime import datetime

class StatisticsSystem:
    def __init__(self):
        self.import_db_file = 'processed_emails_import.db'
        self.export_db_file = 'processed_emails.db'
        
    def init_database(self):
        try:
            # 初始化进口统计数据库
            if os.path.exists(self.import_db_file):
                conn = sqlite3.connect(self.import_db_file)
                cursor = conn.cursor()
                cursor.execute('''
                CREATE TABLE IF NOT EXISTS import_attachments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    attachment_name TEXT NOT NULL,
                    process_date DATE NOT NULL,
                    has_dangerous INTEGER DEFAULT 0,
                    matched_keywords TEXT,
                    sender_email TEXT,
                    subject TEXT,
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(attachment_name, process_date)
                )
                ''')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_att_date ON import_attachments(process_date)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_att_dangerous ON import_attachments(has_dangerous)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_att_keywords ON import_attachments(matched_keywords)')
                conn.commit()
                conn.close()
            
            # 初始化出口统计数据库
            if os.path.exists(self.export_db_file):
                conn = sqlite3.connect(self.export_db_file)
                cursor = conn.cursor()
                cursor.execute('''
                CREATE TABLE IF NOT EXISTS export_attachments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    attachment_name TEXT NOT NULL,
                    process_date DATE NOT NULL,
                    has_dangerous INTEGER DEFAULT 0,
                    matched_keywords TEXT,
                    sender_email TEXT,
                    subject TEXT,
                    created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(attachment_name, process_date)
                )
                ''')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_export_att_date ON export_attachments(process_date)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_export_att_dangerous ON export_attachments(has_dangerous)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_export_att_keywords ON export_attachments(matched_keywords)')
                conn.commit()
                conn.close()
            
            return True
        except Exception as e:
            print(f"初始化失败: {e}")
            return False
        
    def get_date_range(self, db_type):
        try:
            if db_type == 'import':
                db_file = self.import_db_file
                table_name = 'import_attachments'
            else:
                db_file = self.export_db_file
                table_name = 'export_attachments'
            
            if not os.path.exists(db_file):
                return None
            
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            cursor.execute(f'''
            SELECT MIN(process_date), MAX(process_date), COUNT(DISTINCT process_date)
            FROM {table_name}
            ''')
            min_date, max_date, days_count = cursor.fetchone()
            conn.close()
            
            if min_date and max_date:
                return {'min_date': min_date, 'max_date': max_date, 'days_count': days_count or 0}
            return None
        except:
            return None
    
    def get_keywords_summary(self, start_date=None, end_date=None):
        """获取关键词统计摘要（进口和出口分开）"""
        try:
            results = {
                'import': {},
                'export': {},
                'total_import': 0,
                'total_export': 0,
                'all_keywords': set()
            }
            
            # 统计进口关键词
            if os.path.exists(self.import_db_file):
                conn = sqlite3.connect(self.import_db_file)
                cursor = conn.cursor()
                
                if start_date and end_date:
                    cursor.execute('''
                    SELECT matched_keywords, COUNT(*) as count
                    FROM import_attachments
                    WHERE process_date >= ? AND process_date <= ?
                    GROUP BY matched_keywords
                    ORDER BY count DESC
                    ''', (start_date, end_date))
                else:
                    cursor.execute('''
                    SELECT matched_keywords, COUNT(*) as count
                    FROM import_attachments
                    GROUP BY matched_keywords
                    ORDER BY count DESC
                    ''')
                
                for keyword_str, count in cursor.fetchall():
                    if keyword_str:
                        # 关键词可能是逗号分隔的多个
                        keywords = [k.strip() for k in keyword_str.split(',') if k.strip()]
                        for keyword in keywords:
                            results['import'][keyword] = results['import'].get(keyword, 0) + 1
                            results['all_keywords'].add(keyword)
                
                # 获取进口总数
                if start_date and end_date:
                    cursor.execute('SELECT COUNT(*) FROM import_attachments WHERE process_date >= ? AND process_date <= ?', 
                                 (start_date, end_date))
                else:
                    cursor.execute('SELECT COUNT(*) FROM import_attachments')
                results['total_import'] = cursor.fetchone()[0] or 0
                
                conn.close()
            
            # 统计出口关键词
            if os.path.exists(self.export_db_file):
                conn = sqlite3.connect(self.export_db_file)
                cursor = conn.cursor()
                
                if start_date and end_date:
                    cursor.execute('''
                    SELECT matched_keywords, COUNT(*) as count
                    FROM export_attachments
                    WHERE process_date >= ? AND process_date <= ?
                    GROUP BY matched_keywords
                    ORDER BY count DESC
                    ''', (start_date, end_date))
                else:
                    cursor.execute('''
                    SELECT matched_keywords, COUNT(*) as count
                    FROM export_attachments
                    GROUP BY matched_keywords
                    ORDER BY count DESC
                    ''')
                
                for keyword_str, count in cursor.fetchall():
                    if keyword_str:
                        # 关键词可能是逗号分隔的多个
                        keywords = [k.strip() for k in keyword_str.split(',') if k.strip()]
                        for keyword in keywords:
                            results['export'][keyword] = results['export'].get(keyword, 0) + 1
                            results['all_keywords'].add(keyword)
                
                # 获取出口总数
                if start_date and end_date:
                    cursor.execute('SELECT COUNT(*) FROM export_attachments WHERE process_date >= ? AND process_date <= ?', 
                                 (start_date, end_date))
                else:
                    cursor.execute('SELECT COUNT(*) FROM export_attachments')
                results['total_export'] = cursor.fetchone()[0] or 0
                
                conn.close()
            
            # 转换为列表并排序
            results['all_keywords'] = sorted(list(results['all_keywords']))
            
            return results
        except Exception as e:
            print(f"获取关键词摘要失败: {e}")
            return None
            
    def query_statistics_with_keywords(self, db_type, start_date, end_date, selected_keywords=None):
        """根据关键词筛选查询统计"""
        try:
            if db_type == 'import':
                db_file = self.import_db_file
                table_name = 'import_attachments'
            else:
                db_file = self.export_db_file
                table_name = 'export_attachments'
            
            if not os.path.exists(db_file):
                return None
            
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # 基础查询
            query_base = f'''
            SELECT id, attachment_name, process_date,
                   CASE WHEN has_dangerous = 1 THEN '是' ELSE '否' END as has_dangerous,
                   matched_keywords, sender_email, subject,
                   strftime('%H:%M:%S', created_time) as created_time
            FROM {table_name}
            WHERE process_date >= ? AND process_date <= ?
            '''
            
            query_params = [start_date, end_date]
            
            # 如果有选中的关键词，添加关键词筛选
            if selected_keywords:
                keyword_conditions = []
                for keyword in selected_keywords:
                    keyword_conditions.append(f"matched_keywords LIKE ?")
                    query_params.append(f'%{keyword}%')
                
                query_base += f" AND ({' OR '.join(keyword_conditions)})"
            
            query_base += " ORDER BY process_date DESC, created_time DESC"
            
            cursor.execute(query_base, query_params)
            attachments = cursor.fetchall()
            
            # 统计总数
            count_query = f'''
            SELECT COUNT(*), SUM(CASE WHEN has_dangerous = 1 THEN 1 ELSE 0 END),
                   SUM(CASE WHEN has_dangerous = 0 THEN 1 ELSE 0 END)
            FROM {table_name}
            WHERE process_date >= ? AND process_date <= ?
            '''
            
            count_params = [start_date, end_date]
            
            if selected_keywords:
                keyword_conditions = []
                for keyword in selected_keywords:
                    keyword_conditions.append(f"matched_keywords LIKE ?")
                    count_params.append(f'%{keyword}%')
                
                count_query += f" AND ({' OR '.join(keyword_conditions)})"
            
            cursor.execute(count_query, count_params)
            total, dangerous, non_dangerous = cursor.fetchone()
            conn.close()
            
            return {
                'db_type': db_type,
                'start_date': start_date,
                'end_date': end_date,
                'total_attachments': total or 0,
                'dangerous_attachments': dangerous or 0,
                'non_dangerous_attachments': non_dangerous or 0,
                'attachments': attachments,
                'selected_keywords': selected_keywords or [],
                'query_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
        except Exception as e:
            print(f"查询失败: {e}")
            return None
            
    def query_statistics(self, db_type, start_date, end_date, use_cache=False):
        """原始查询方法（保持兼容性）"""
        return self.query_statistics_with_keywords(db_type, start_date, end_date, None)
            
    def delete_attachments(self, db_type, attachment_ids):
        try:
            if db_type == 'import':
                db_file = self.import_db_file
                table_name = 'import_attachments'
            else:
                db_file = self.export_db_file
                table_name = 'export_attachments'
            
            if not os.path.exists(db_file) or not attachment_ids:
                return 0
            
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            placeholders = ','.join(['?'] * len(attachment_ids))
            
            cursor.execute(f'''
            DELETE FROM {table_name} WHERE id IN ({placeholders})
            ''', attachment_ids)
            
            deleted_count = cursor.rowcount
            conn.commit()
            conn.close()
            return deleted_count
        except:
            return 0
            
    def export_statistics(self, db_type, start_date, end_date, export_format='csv', selected_keywords=None):
        try:
            result = self.query_statistics_with_keywords(db_type, start_date, end_date, selected_keywords)
            if not result:
                return None
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            if export_format.lower() == 'csv':
                filename = f'{db_type}_statistics_{start_date}_to_{end_date}_{timestamp}.csv'
                with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['统计类型', '开始日期', '结束日期', '查询时间', '筛选关键词'])
                    keywords_str = ','.join(selected_keywords) if selected_keywords else '全部'
                    writer.writerow([db_type, start_date, end_date, result['query_time'], keywords_str])
                    writer.writerow([])
                    writer.writerow(['统计摘要'])
                    writer.writerow(['总附件数', '含危险品附件数', '不含危险品附件数'])
                    writer.writerow([result['total_attachments'], 
                                   result['dangerous_attachments'], 
                                   result['non_dangerous_attachments']])
                    writer.writerow([])
                    writer.writerow(['详细附件记录'])
                    writer.writerow(['ID', '附件名称', '处理日期', '是否含危险品', 
                                   '匹配关键词', '发件人邮箱', '邮件主题', '创建时间'])
                    for attachment in result['attachments']:
                        writer.writerow(attachment)
                return filename
            return None
        except:
            return None
            
    def clear_cache(self, db_type=None):
        return True
    
    def add_attachment_record(self, db_type, attachment_info):
        """添加附件统计记录"""
        try:
            if db_type == 'import':
                db_file = self.import_db_file
                table_name = 'import_attachments'
            else:
                db_file = self.export_db_file
                table_name = 'export_attachments'
            
            # 确保数据库文件存在，如果不存在则创建
            if not os.path.exists(db_file):
                print(f"⚠️ 数据库文件 {db_file} 不存在，正在创建...")
                self.init_database()
                if not os.path.exists(db_file):
                    print(f"❌ 无法创建数据库文件: {db_file}")
                    return False
            
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # 确保表存在
            cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS {table_name} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                attachment_name TEXT NOT NULL,
                process_date DATE NOT NULL,
                has_dangerous INTEGER DEFAULT 0,
                matched_keywords TEXT,
                sender_email TEXT,
                subject TEXT,
                created_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(attachment_name, process_date)
            )
            ''')
            
            # 准备数据
            attachment_name = attachment_info.get('attachment_name', '')
            process_date = attachment_info.get('process_date', datetime.now().strftime('%Y-%m-%d'))
            has_dangerous = attachment_info.get('has_dangerous', 0)
            matched_keywords = attachment_info.get('matched_keywords', '')
            sender_email = attachment_info.get('sender_email', '')
            subject = attachment_info.get('subject', '')
            
            if not attachment_name:
                print("❌ 附件名称为空，无法添加统计记录")
                conn.close()
                return False
            
            # 插入记录，如果已经存在则忽略
            cursor.execute(f'''
            INSERT OR IGNORE INTO {table_name} 
            (attachment_name, process_date, has_dangerous, matched_keywords, sender_email, subject)
            VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                attachment_name,
                process_date,
                has_dangerous,
                matched_keywords,
                sender_email,
                subject
            ))
            
            conn.commit()
            rowcount = cursor.rowcount
            conn.close()
            
            if rowcount > 0:
                print(f"✅ 成功添加统计记录: {attachment_name}")
                return True
            else:
                print(f"⚠️ 统计记录已存在: {attachment_name}")
                return False
                
        except Exception as e:
            print(f"❌ 添加附件统计记录失败: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def get_all_keywords(self):
        """获取所有出现过的关键词"""
        try:
            all_keywords = set()
            
            # 从进口数据库获取
            if os.path.exists(self.import_db_file):
                conn = sqlite3.connect(self.import_db_file)
                cursor = conn.cursor()
                cursor.execute('SELECT DISTINCT matched_keywords FROM import_attachments WHERE matched_keywords IS NOT NULL AND matched_keywords != ""')
                for row in cursor.fetchall():
                    if row[0]:
                        keywords = [k.strip() for k in row[0].split(',') if k.strip()]
                        all_keywords.update(keywords)
                conn.close()
            
            # 从出口数据库获取
            if os.path.exists(self.export_db_file):
                conn = sqlite3.connect(self.export_db_file)
                cursor = conn.cursor()
                cursor.execute('SELECT DISTINCT matched_keywords FROM export_attachments WHERE matched_keywords IS NOT NULL AND matched_keywords != ""')
                for row in cursor.fetchall():
                    if row[0]:
                        keywords = [k.strip() for k in row[0].split(',') if k.strip()]
                        all_keywords.update(keywords)
                conn.close()
            
            return sorted(list(all_keywords))
        except Exception as e:
            print(f"获取所有关键词失败: {e}")
            return []