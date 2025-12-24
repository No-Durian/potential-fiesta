"""
安装统计功能所需的依赖
"""

import subprocess
import sys
import os
import platform

def install_package(package):
    """安装指定的Python包"""
    try:
        print(f"正在安装 {package}...")
        # 使用国内镜像源加速安装
        subprocess.check_call([sys.executable, "-m", "pip", "install", 
                             "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", package])
        print(f"✅ {package} 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {package} 安装失败: {e}")
        # 尝试使用默认源
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"✅ {package} 安装成功")
            return True
        except:
            return False

def main():
    """主函数"""
    print("=" * 60)
    print("舱单附件统计系统 - 依赖安装工具")
    print("=" * 60)
    print(f"操作系统: {platform.system()} {platform.release()}")
    print(f"Python版本: {sys.version}")
    print()
    
    required_packages = [
        'tkcalendar',  # 日期选择控件
        'pandas',      # 数据处理
        'openpyxl',    # Excel支持
    ]
    
    print("需要安装以下包:")
    for package in required_packages:
        print(f"  - {package}")
    
    print("\n开始安装...")
    
    success_count = 0
    for package in required_packages:
        if install_package(package):
            success_count += 1
    
    print("\n" + "=" * 60)
    if success_count == len(required_packages):
        print("✅ 所有依赖安装成功！")
        print("您现在可以使用完整的统计查询功能了。")
    else:
        print(f"⚠️  成功安装了 {success_count}/{len(required_packages)} 个包")
        print("某些功能可能无法正常使用。")
    
    print("=" * 60)
    
    # 初始化统计数据库
    try:
        print("\n正在初始化统计数据库...")
        sys.path.append(os.path.dirname(os.path.abspath(__file__)))
        
        # 创建统计系统模块
        stats_code = '''
import sqlite3
import os
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
                cursor.execute(\'''
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
                \''')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_att_date ON import_attachments(process_date)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_import_att_dangerous ON import_attachments(has_dangerous)')
                conn.commit()
                conn.close()
            
            # 初始化出口统计数据库
            if os.path.exists(self.export_db_file):
                conn = sqlite3.connect(self.export_db_file)
                cursor = conn.cursor()
                cursor.execute(\'''
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
                \''')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_export_att_date ON export_attachments(process_date)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_export_att_dangerous ON export_attachments(has_dangerous)')
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
            cursor.execute(f\'''
            SELECT MIN(process_date), MAX(process_date), COUNT(DISTINCT process_date)
            FROM {table_name}
            \''')
            min_date, max_date, days_count = cursor.fetchone()
            conn.close()
            
            if min_date and max_date:
                return {'min_date': min_date, 'max_date': max_date, 'days_count': days_count or 0}
            return None
        except:
            return None
            
    def query_statistics(self, db_type, start_date, end_date, use_cache=False):
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
            
            # 查询详细附件信息
            cursor.execute(f\'''
            SELECT id, attachment_name, process_date,
                   CASE WHEN has_dangerous = 1 THEN '是' ELSE '否' END as has_dangerous,
                   matched_keywords, sender_email, subject,
                   strftime('%H:%M:%S', created_time) as created_time
            FROM {table_name}
            WHERE process_date >= ? AND process_date <= ?
            ORDER BY process_date DESC, created_time DESC
            \''', (start_date, end_date))
            
            attachments = cursor.fetchall()
            
            # 统计总数
            cursor.execute(f\'''
            SELECT COUNT(*), SUM(CASE WHEN has_dangerous = 1 THEN 1 ELSE 0 END),
                   SUM(CASE WHEN has_dangerous = 0 THEN 1 ELSE 0 END)
            FROM {table_name}
            WHERE process_date >= ? AND process_date <= ?
            \''', (start_date, end_date))
            
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
                'query_time': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
        except Exception as e:
            print(f"查询失败: {e}")
            return None
            
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
            
            cursor.execute(f\'''
            DELETE FROM {table_name} WHERE id IN ({placeholders})
            \''', attachment_ids)
            
            deleted_count = cursor.rowcount
            conn.commit()
            conn.close()
            return deleted_count
        except:
            return 0
            
    def export_statistics(self, db_type, start_date, end_date, export_format='csv'):
        try:
            result = self.query_statistics(db_type, start_date, end_date)
            if not result:
                return None
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            if export_format.lower() == 'csv':
                filename = f'{db_type}_statistics_{start_date}_to_{end_date}_{timestamp}.csv'
                with open(filename, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['统计类型', '开始日期', '结束日期', '查询时间'])
                    writer.writerow([db_type, start_date, end_date, result['query_time']])
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
'''
        
        # 保存统计系统模块
        with open('statistics_system.py', 'w', encoding='utf-8') as f:
            f.write(stats_code)
        
        print("✅ 统计系统模块已创建")
        
        # 测试初始化
        from statistics_system import StatisticsSystem
        stats = StatisticsSystem()
        if stats.init_database():
            print("✅ 统计数据库初始化成功")
        else:
            print("❌ 统计数据库初始化失败")
            
    except Exception as e:
        print(f"❌ 初始化统计系统时出错: {e}")
    
    print("\n安装完成！您现在可以:")
    print("1. 运行 'python SimpleAutoRW_GUI_with_stats.py' 启动完整系统")
    print("2. 在GUI左侧使用统计查询功能")
    print("3. 在'统计功能'菜单中初始化统计数据库")

if __name__ == "__main__":
    main()
