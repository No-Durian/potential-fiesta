# test_email_send.py
"""
测试邮件发送功能
"""

import sys
import os
import datetime  # 添加这行
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config_manager import ConfigManager

def test_smtp_connection():
    """测试SMTP连接"""
    try:
        import smtplib
        import ssl
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        
        cm = ConfigManager()
        email_config = cm.get_email_config()
        
        print("="*60)
        print("测试SMTP连接")
        print("="*60)
        
        # 测试进口邮箱SMTP
        print(f"\n1. 测试进口邮箱SMTP连接 ({email_config['smtp_server']}:{email_config['smtp_port']})")
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(
                email_config['smtp_server'], 
                email_config['smtp_port'],
                context=context
            ) as server:
                server.login(email_config['import_email'], email_config['import_password'])
                print("✅ 进口邮箱SMTP连接成功")
                
                # 发送测试邮件
                msg = MIMEMultipart()
                msg['From'] = email_config['import_email']
                msg['To'] = email_config['import_email']  # 发给自己
                msg['Subject'] = "测试邮件 - 进口舱单系统"
                
                body = f"这是一封测试邮件，用于验证SMTP连接和发送功能。\n\n系统时间：{datetime.datetime.now()}"
                msg.attach(MIMEText(body, 'plain', 'utf-8'))
                
                server.send_message(msg)
                print("✅ 测试邮件发送成功")
                
        except Exception as e:
            print(f"❌ 进口邮箱SMTP连接失败: {e}")
        
        # 测试出口邮箱SMTP
        print(f"\n2. 测试出口邮箱SMTP连接 ({email_config['smtp_server']}:{email_config['smtp_port']})")
        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(
                email_config['smtp_server'], 
                email_config['smtp_port'],
                context=context
            ) as server:
                server.login(email_config['export_email'], email_config['export_password'])
                print("✅ 出口邮箱SMTP连接成功")
                
                # 发送测试邮件
                msg = MIMEMultipart()
                msg['From'] = email_config['export_email']
                msg['To'] = email_config['export_email']  # 发给自己
                msg['Subject'] = "测试邮件 - 出口舱单系统"
                
                body = f"这是一封测试邮件，用于验证SMTP连接和发送功能。\n\n系统时间：{datetime.datetime.now()}"
                msg.attach(MIMEText(body, 'plain', 'utf-8'))
                
                server.send_message(msg)
                print("✅ 测试邮件发送成功")
                
        except Exception as e:
            print(f"❌ 出口邮箱SMTP连接失败: {e}")
            
    except Exception as e:
        print(f"❌ SMTP测试失败: {e}")

def test_pop3_connection():
    """测试POP3连接"""
    try:
        import poplib
        
        cm = ConfigManager()
        email_config = cm.get_email_config()
        
        print("\n" + "="*60)
        print("测试POP3连接")
        print("="*60)
        
        # 测试进口邮箱POP3
        print(f"\n1. 测试进口邮箱POP3连接 ({email_config['pop3_server']}:{email_config['pop3_port']})")
        try:
            server = poplib.POP3_SSL(email_config['pop3_server'], email_config['pop3_port'])
            server.user(email_config['import_email'])
            server.pass_(email_config['import_password'])
            
            # 获取邮件统计
            count, size = server.stat()
            print(f"✅ 进口邮箱POP3连接成功")
            print(f"   邮件数量: {count}")
            print(f"   总大小: {size} 字节")
            
            server.quit()
            
        except Exception as e:
            print(f"❌ 进口邮箱POP3连接失败: {e}")
        
    except Exception as e:
        print(f"❌ POP3测试失败: {e}")

def main():
    """主测试函数"""
    print("邮件系统功能测试")
    print("="*60)
    
    # 检查配置
    cm = ConfigManager()
    configs = cm.get_all_configs()
    
    print("\n1. 检查配置:")
    for section, values in configs.items():
        print(f"\n{section}:")
        for key, value in values.items():
            if 'password' in key:
                value = '*' * len(str(value)) if value else '空'
            print(f"  {key}: {value}")
    
    # 测试连接
    test_smtp_connection()
    test_pop3_connection()
    
    print("\n" + "="*60)
    print("测试完成")
    print("="*60)

if __name__ == "__main__":
    main()