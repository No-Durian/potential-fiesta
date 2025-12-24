# check_system.py
"""
系统诊断脚本
"""

import os
import sys
import sqlite3
import configparser
import logging

def setup_logging():
    """设置诊断日志"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('system_diagnostic.log', encoding='utf-8')
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

def check_config():
    """检查配置文件"""
    print("\n" + "="*60)
    print("检查配置文件 config.ini")
    print("="*60)
    
    if not os.path.exists('config.ini'):
        print("❌ config.ini 不存在")
        return False
    
    config = configparser.ConfigParser()
    config.optionxform = str  # 保留大小写
    config.read('config.ini', encoding='utf-8')
    
    print("配置文件内容:")
    for section in config.sections():
        print(f"\n[{section}]")
        for key, value in config.items(section):
            if 'password' in key.lower():
                value = '*' * len(value)
            print(f"  {key} = {value[:50]}{'...' if len(value) > 50 else ''}")
    
    return True

def check_database():
    """检查数据库"""
    print("\n" + "="*60)
    print("检查数据库")
    print("="*60)
    
    databases = [
        ('processed_emails_import.db', '进口'),
        ('processed_emails.db', '出口')
    ]
    
    for db_file, db_type in databases:
        print(f"\n{db_type}数据库 ({db_file}):")
        if not os.path.exists(db_file):
            print("  ❌ 文件不存在")
            continue
        
        try:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            
            # 检查表
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = cursor.fetchall()
            print(f"  ✅ 文件大小: {os.path.getsize(db_file)} 字节")
            print(f"  表: {[t[0] for t in tables]}")
            
            if 'keyword_emails' in [t[0] for t in tables]:
                # 检查列
                cursor.execute("PRAGMA table_info(keyword_emails)")
                columns = cursor.fetchall()
                print(f"  列结构:")
                for col in columns:
                    print(f"    {col[1]} ({col[2]})")
                
                # 查看数据
                cursor.execute("SELECT COUNT(*) FROM keyword_emails")
                count = cursor.fetchone()[0]
                print(f"  记录数: {count}")
                
                if count > 0:
                    cursor.execute("SELECT matched_keywords FROM keyword_emails LIMIT 5")
                    keywords = cursor.fetchall()
                    print(f"  关键词示例: {[k[0] for k in keywords]}")
            
            conn.close()
            
        except Exception as e:
            print(f"  ❌ 检查失败: {e}")

def check_keywords():
    """检查关键词配置"""
    print("\n" + "="*60)
    print("检查关键词配置")
    print("="*60)
    
    try:
        from config_manager import ConfigManager
        cm = ConfigManager()
        
        # 获取所有关键词
        keywords = cm.get_keywords()
        print(f"进口关键词: {keywords.get('import', [])}")
        print(f"出口关键词: {keywords.get('export', [])}")
        
        # 获取翻译映射
        translation = cm.get_keyword_translation_map()
        print(f"\n关键词翻译映射 ({len(translation)} 个):")
        for en, cn in translation.items():
            print(f"  {en} -> {cn}")
        
        return True
        
    except Exception as e:
        print(f"❌ 检查关键词失败: {e}")
        return False

def check_modules():
    """检查模块依赖"""
    print("\n" + "="*60)
    print("检查模块依赖")
    print("="*60)
    
    required_modules = [
        'poplib', 'email', 'smtplib', 'sqlite3',
        'configparser', 'logging', 'threading'
    ]
    
    for module in required_modules:
        try:
            __import__(module)
            print(f"✅ {module}")
        except ImportError as e:
            print(f"❌ {module}: {e}")

def test_email_config():
    """测试邮件配置"""
    print("\n" + "="*60)
    print("测试邮件配置")
    print("="*60)
    
    try:
        from config_manager import ConfigManager
        cm = ConfigManager()
        
        email_config = cm.get_email_config()
        print("邮件配置:")
        for key, value in email_config.items():
            if 'password' in key:
                value = '*' * len(value) if value else '空'
            print(f"  {key}: {value}")
        
        # 检查必要配置是否为空
        required = ['import_email', 'import_password', 'export_email', 'export_password']
        missing = [key for key in required if not email_config.get(key)]
        
        if missing:
            print(f"\n❌ 缺失必要配置: {missing}")
            return False
        else:
            print("\n✅ 邮件配置完整")
            return True
            
    except Exception as e:
        print(f"❌ 检查邮件配置失败: {e}")
        return False

def main():
    """主诊断函数"""
    print("="*60)
    print("舱单邮件处理系统 - 诊断工具")
    print("="*60)
    
    # 检查当前目录
    print(f"当前目录: {os.getcwd()}")
    print(f"Python版本: {sys.version}")
    
    # 执行各项检查
    check_config()
    check_database()
    check_keywords()
    check_modules()
    test_email_config()
    
    print("\n" + "="*60)
    print("诊断完成")
    print("="*60)

if __name__ == "__main__":
    main()