"""
历史邮件同步检查脚本
用于验证数据库同步情况
"""

import sqlite3
import os
from datetime import datetime

def check_column_exists(db_file, table_name, column_name):
    """检查表中是否存在指定列"""
    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [info[1] for info in cursor.fetchall()]
        
        conn.close()
        return column_name in columns
    except Exception as e:
        print(f"❌ 检查列失败: {e}")
        return False

def get_synced_count(db_file):
    """获取同步的邮件数量"""
    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 先检查表是否存在
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails'")
        if not cursor.fetchone():
            return 0, 0, False
        
        # 检查sync_source列是否存在
        has_sync_source = check_column_exists(db_file, 'keyword_emails', 'sync_source')
        
        # 获取总邮件数
        cursor.execute("SELECT COUNT(*) FROM keyword_emails")
        total = cursor.fetchone()[0]
        
        # 获取同步邮件数
        if has_sync_source:
            cursor.execute("SELECT COUNT(*) FROM keyword_emails WHERE sync_source = 'history_sync'")
            synced = cursor.fetchone()[0]
        else:
            synced = 0
        
        conn.close()
        return total, synced, has_sync_source
        
    except Exception as e:
        print(f"❌ 获取统计失败: {e}")
        return 0, 0, False

def check_sync_status():
    """检查同步状态"""
    
    print("=" * 60)
    print("历史邮件同步状态检查")
    print("=" * 60)
    
    # 检查进口数据库
    if os.path.exists('processed_emails_import.db'):
        total, synced, has_sync_source = get_synced_count('processed_emails_import.db')
        
        print(f"进口数据库:")
        print(f"  总邮件数: {total}")
        if has_sync_source:
            print(f"  历史同步: {synced}")
            print(f"  系统处理: {total - synced}")
        else:
            print(f"  历史同步: 0 (表结构需要升级)")
            print(f"  ⚠️  请运行 UpdateDatabaseSchema.py 升级数据库")
    else:
        print("进口数据库: 文件不存在")
    
    print()
    
    # 检查出口数据库
    if os.path.exists('processed_emails.db'):
        total, synced, has_sync_source = get_synced_count('processed_emails.db')
        
        print(f"出口数据库:")
        print(f"  总邮件数: {total}")
        if has_sync_source:
            print(f"  历史同步: {synced}")
            print(f"  系统处理: {total - synced}")
        else:
            print(f"  历史同步: 0 (表结构需要升级)")
            print(f"  ⚠️  请运行 UpdateDatabaseSchema.py 升级数据库")
    else:
        print("出口数据库: 文件不存在")
    
    print("=" * 60)

def show_recent_synced_mails(limit=10):
    """显示最近同步的邮件"""
    
    print("\n最近同步的历史邮件:")
    print("-" * 60)
    
    databases = [
        ('processed_emails_import.db', '进口'),
        ('processed_emails.db', '出口')
    ]
    
    for db_file, db_type in databases:
        if os.path.exists(db_file):
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # 先检查表是否存在
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails'")
            if not cursor.fetchone():
                print(f"\n{db_type}数据库: 表不存在")
                conn.close()
                continue
            
            # 检查是否有sync_source列
            has_sync_source = check_column_exists(db_file, 'keyword_emails', 'sync_source')
            
            if has_sync_source:
                cursor.execute('''
                SELECT COUNT(*) FROM keyword_emails WHERE sync_source = 'history_sync'
                ''')
                count = cursor.fetchone()[0]
                
                if count > 0:
                    print(f"\n{db_type}数据库 ({count} 封同步邮件):")
                    
                    cursor.execute('''
                    SELECT processed_date, sender, subject, matched_keywords
                    FROM keyword_emails 
                    WHERE sync_source = 'history_sync'
                    ORDER BY processed_date DESC
                    LIMIT ?
                    ''', (limit,))
                    
                    for row in cursor.fetchall():
                        print(f"  时间: {row[0]}")
                        print(f"  发件人: {row[1]}")
                        print(f"  主题: {row[2][:50]}...")
                        print(f"  关键词: {row[3]}")
                        print()
                else:
                    print(f"\n{db_type}数据库: 没有同步的历史邮件")
            else:
                print(f"\n{db_type}数据库: 表结构需要升级，无法显示同步邮件")
            
            conn.close()

def check_database_health():
    """检查数据库健康状况"""
    print("\n数据库健康检查:")
    print("-" * 40)
    
    for db_file, db_name in [('processed_emails_import.db', '进口数据库'), 
                             ('processed_emails.db', '出口数据库')]:
        if os.path.exists(db_file):
            try:
                conn = sqlite3.connect(db_file)
                cursor = conn.cursor()
                
                # 检查表
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
                tables = [table[0] for table in cursor.fetchall()]
                
                print(f"\n{db_name}:")
                print(f"  文件大小: {os.path.getsize(db_file) / 1024:.1f} KB")
                print(f"  表数量: {len(tables)}")
                
                if 'keyword_emails' in tables:
                    cursor.execute("SELECT COUNT(*) FROM keyword_emails")
                    count = cursor.fetchone()[0]
                    print(f"  keyword_emails表记录数: {count}")
                    
                    # 检查列
                    cursor.execute("PRAGMA table_info(keyword_emails)")
                    columns = cursor.fetchall()
                    print(f"  keyword_emails表列数: {len(columns)}")
                    
                    # 检查是否有sync_source列
                    column_names = [col[1] for col in columns]
                    if 'sync_source' in column_names:
                        print(f"  ✅ 包含sync_source列")
                    else:
                        print(f"  ❌ 缺少sync_source列")
                
                conn.close()
                
            except Exception as e:
                print(f"❌ 检查{db_name}失败: {e}")
        else:
            print(f"\n{db_name}: 文件不存在")

if __name__ == "__main__":
    check_database_health()
    check_sync_status()
    show_recent_synced_mails(5)
    
    print("\n" + "=" * 60)
    print("建议操作:")
    print("1. 如果缺少sync_source列，请运行: python UpdateDatabaseSchema.py")
    print("2. 同步历史邮件，请运行: python HistoryMailSync.py")
    print("3. 或在GUI中点击'同步历史邮件'按钮")
    print("=" * 60)