"""
数据库表结构升级脚本
用于添加sync_source列和其他必要的字段
"""

import sqlite3
import os

def check_and_add_column(db_file, table_name, column_name, column_type="TEXT"):
    """
    检查并添加列到表中
    """
    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 检查表是否存在
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        if not cursor.fetchone():
            print(f"❌ 表 {table_name} 不存在于 {db_file}")
            conn.close()
            return False
        
        # 检查列是否存在
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [info[1] for info in cursor.fetchall()]
        
        if column_name in columns:
            print(f"✅ {db_file} 中的 {table_name} 表已存在 {column_name} 列")
            conn.close()
            return True
        else:
            # 添加列
            print(f"🔧 正在为 {db_file} 中的 {table_name} 表添加 {column_name} 列...")
            cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type} DEFAULT ''")
            conn.commit()
            print(f"✅ 成功添加 {column_name} 列到 {db_file}")
            conn.close()
            return True
            
    except Exception as e:
        print(f"❌ 更新 {db_file} 失败: {e}")
        return False

def upgrade_import_db():
    """升级进口数据库"""
    if not os.path.exists('processed_emails_import.db'):
        print("❌ 进口数据库文件不存在")
        return False
    
    print("=" * 60)
    print("升级进口数据库表结构")
    print("=" * 60)
    
    success = check_and_add_column('processed_emails_import.db', 'keyword_emails', 'sync_source')
    
    if success:
        print("✅ 进口数据库升级完成")
    else:
        print("❌ 进口数据库升级失败")
    
    return success

def upgrade_export_db():
    """升级出口数据库"""
    if not os.path.exists('processed_emails.db'):
        print("❌ 出口数据库文件不存在")
        return False
    
    print("\n" + "=" * 60)
    print("升级出口数据库表结构")
    print("=" * 60)
    
    success = check_and_add_column('processed_emails.db', 'keyword_emails', 'sync_source')
    
    if success:
        print("✅ 出口数据库升级完成")
    else:
        print("❌ 出口数据库升级失败")
    
    return success

def check_table_structure(db_file, table_name):
    """检查表结构"""
    try:
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = cursor.fetchall()
        
        print(f"\n{table_name} 表结构:")
        print("-" * 40)
        for col in columns:
            print(f"  {col[1]} ({col[2]})")
        
        conn.close()
        return True
    except Exception as e:
        print(f"❌ 检查表结构失败: {e}")
        return False

def main():
    """主函数"""
    print("数据库表结构升级工具")
    print("=" * 60)
    
    # 升级进口数据库
    import_success = upgrade_import_db()
    
    # 升级出口数据库
    export_success = upgrade_export_db()
    
    # 显示表结构
    if os.path.exists('processed_emails_import.db'):
        check_table_structure('processed_emails_import.db', 'keyword_emails')
    
    if os.path.exists('processed_emails.db'):
        check_table_structure('processed_emails.db', 'keyword_emails')
    
    print("\n" + "=" * 60)
    if import_success and export_success:
        print("✅ 数据库升级完成！")
    else:
        print("⚠️  数据库升级完成，但可能存在一些问题")
    print("=" * 60)

if __name__ == "__main__":
    main()