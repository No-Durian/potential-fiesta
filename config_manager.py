# config_manager.py
import configparser
import os
import json
from datetime import datetime
import logging

# 设置日志记录器
def setup_logger(name):
    """设置日志记录器"""
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
    return logger

class ConfigManager:
    def __init__(self, config_path='config.ini'):
        self.config_path = config_path
        # 保留配置项大小写（尤其是 keyword_translation 的英文关键词）
        self.config = configparser.ConfigParser()
        self.config.optionxform = str
        # 初始化日志记录器
        self.logger = setup_logger('ConfigManager')
        self.load_config()
        
        # 确保所有配置节存在
        self._ensure_sections()
    
    def _ensure_sections(self):
        """确保所有必要的配置节存在"""
        required_sections = [
            'email', 'keywords', 'keyword_translation', 'sms', 'files',
            'settings', 'web', 'additional_recipients'
        ]
        for section in required_sections:
            if not self.config.has_section(section):
                self.config.add_section(section)
                self.logger.info(f"创建配置节: {section}")
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_path):
            try:
                self.config.read(self.config_path, encoding='utf-8')
                self.logger.info(f"成功加载配置文件: {self.config_path}")
            except Exception as e:
                self.logger.error(f"加载配置文件失败: {e}")
                # 创建默认配置
                self._create_default_config()
        else:
            self.logger.warning(f"配置文件不存在: {self.config_path}")
            # 创建默认配置
            self._create_default_config()
            self.save_config()
        return self.config
    
    def _create_default_config(self):
        """创建默认配置"""
        self.logger.info("创建默认配置")
        
        # 邮箱配置默认值
        self.config.set('email', '进口邮箱地址', '')
        self.config.set('email', '进口邮箱密码', '')
        self.config.set('email', '出口邮箱地址', '')
        self.config.set('email', '出口密码', '')
        self.config.set('email', 'pop3服务器', 'pop.qq.com')
        self.config.set('email', 'pop3端口', '995')
        self.config.set('email', 'smtp服务器', 'smtp.qq.com')
        self.config.set('email', 'smtp端口', '465')
        
        # 关键词配置默认值
        self.config.set('keywords', '进口关键词1', 'Calcium Nitrate')
        self.config.set('keywords', '进口关键词2', 'Calcium Nitrate Tetrahydrate')
        self.config.set('keywords', '进口关键词3', 'Magnesium Nitrate Hexahydrate')
        self.config.set('keywords', '出口关键词1', 'Calcium Nitrate')
        self.config.set('keywords', '出口关键词2', 'Calcium Nitrate Tetrahydrate')
        self.config.set('keywords', '出口关键词3', 'Magnesium Nitrate Hexahydrate')
        
        # 短信配置默认值
        self.config.set('sms', '短信账户', '')
        self.config.set('sms', '短信密码', '')
        self.config.set('sms', '短信手机号', '')
        self.config.set('sms', '进口短信模板', '【天津港集装箱码头有限公司】进口舱单处理程序异常退出: {error_msg}')
        self.config.set('sms', '出口短信模板', '【天津港集装箱码头有限公司】出口舱单处理程序异常退出: {error_msg}')
        self.config.set('sms', '短信API地址', '')
        
        # 文件路径配置默认值
        self.config.set('files', '进口数据库', 'processed_emails_import.db')
        self.config.set('files', '出口数据库', 'processed_emails.db')
        self.config.set('files', '进口日志文件', 'email_processing_log_import.csv')
        self.config.set('files', '出口日志文件', 'email_processing_log.csv')
        
        # 系统设置默认值
        self.config.set('settings', '检查间隔', '30')
        # 按需求：日志每 51 天自动清理一次（自动检测只检测最近 50 天邮件）
        self.config.set('settings', '日志保留天数', '51')
        self.config.set('settings', '数据库保留天数', '90')
        self.config.set('settings', '界面主题', 'dark-blue')
        self.config.set('settings', '字体大小', '12')

        # 关键词中英文映射（用于Excel中文货名列）
        # 说明：用户只配置“关键词”，不再额外配置“回复语句”。中文货名映射自动维护。
        self.config.set('keyword_translation', 'Calcium Nitrate', '硝酸钙')
        self.config.set('keyword_translation', 'Calcium Nitrate Tetrahydrate', '四水合硝酸钙')
        self.config.set('keyword_translation', 'Magnesium Nitrate Hexahydrate', '六水合硝酸镁')
        
        # Web配置默认值
        self.config.set('web', '主机', '0.0.0.0')
        self.config.set('web', '端口', '5000')
        self.config.set('web', '调试模式', 'False')
        self.config.set('web', '统计天数', '30')
        self.config.set('web', '图表DPI', '100')
        self.config.set('web', '监控间隔', '30')
        
        # 额外收件人配置默认值（空）
        # 不需要设置默认值，用户自行添加
    
    def save_config(self):
        """保存配置到文件"""
        try:
            # 确保目录存在
            config_dir = os.path.dirname(self.config_path)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir, exist_ok=True)
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                self.config.write(f)
            
            self.logger.info(f"配置已保存到 {self.config_path}")
            return True
        except Exception as e:
            self.logger.error(f"保存配置失败: {e}")
            return False
    
    def update_runtime_config(self):
        """更新运行时配置"""
        try:
            # 重新加载配置
            self.load_config()
            
            # 这里可以添加代码来通知其他模块配置已更新
            # 例如，可以设置一个标志或调用模块的重新初始化方法
            
            self.logger.info("配置已更新，需要重启系统才能生效")
            return True
        except Exception as e:
            self.logger.error(f"更新运行时配置失败: {e}")
            return False
    
    # 邮箱配置
    def get_email_config(self):
        """获取邮箱配置"""
        try:
            return {
                'import_email': self.config.get('email', '进口邮箱地址', fallback=''),
                'import_password': self.config.get('email', '进口邮箱密码', fallback=''),
                'export_email': self.config.get('email', '出口邮箱地址', fallback=''),
                'export_password': self.config.get('email', '出口密码', fallback=''),
                'pop3_server': self.config.get('email', 'pop3服务器', fallback='pop.qq.com'),
                'pop3_port': self.config.getint('email', 'pop3端口', fallback=995),
                'smtp_server': self.config.get('email', 'smtp服务器', fallback='smtp.qq.com'),
                'smtp_port': self.config.getint('email', 'smtp端口', fallback=465)
            }
        except Exception as e:
            self.logger.error(f"获取邮箱配置失败: {e}")
            return {}
    
    def set_email_config(self, email_config):
        """保存邮箱配置"""
        try:
            # 验证必要字段
            required_fields = ['import_email', 'import_password', 'export_email', 'export_password']
            for field in required_fields:
                if field not in email_config or not email_config[field]:
                    raise ValueError(f"缺少必要字段: {field}")
            
            for key, value in email_config.items():
                if key == 'import_email':
                    self.config.set('email', '进口邮箱地址', str(value))
                elif key == 'import_password':
                    self.config.set('email', '进口邮箱密码', str(value))
                elif key == 'export_email':
                    self.config.set('email', '出口邮箱地址', str(value))
                elif key == 'export_password':
                    self.config.set('email', '出口密码', str(value))
                elif key == 'pop3_server':
                    self.config.set('email', 'pop3服务器', str(value))
                elif key == 'pop3_port':
                    self.config.set('email', 'pop3端口', str(value))
                elif key == 'smtp_server':
                    self.config.set('email', 'smtp服务器', str(value))
                elif key == 'smtp_port':
                    self.config.set('email', 'smtp端口', str(value))
            
            # 保存配置
            if self.save_config():
                self.update_runtime_config()
                self.logger.info("邮箱配置已保存")
                return True
            else:
                self.logger.error("保存邮箱配置失败")
                return False
                
        except Exception as e:
            self.logger.error(f"保存邮箱配置失败: {e}")
            return False
    
    # 关键词配置
# 修改 config_manager.py 中的 get_keywords 方法：

    def get_keywords(self):
        """获取关键词配置"""
        import_keywords = []
        export_keywords = []
        
        try:
            # 确保 keywords 节存在
            if not self.config.has_section('keywords'):
                self.config.add_section('keywords')
                self.logger.warning("keywords 节不存在，已创建")
            
            # 先尝试从配置读取
            for key in self.config['keywords']:
                value = self.config['keywords'][key].strip()
                if value:  # 只添加非空值
                    if key.startswith('进口关键词'):
                        import_keywords.append(value)
                    elif key.startswith('出口关键词'):
                        export_keywords.append(value)
            
            # 如果没有读取到关键词，检查是否有默认值
            if not import_keywords and not export_keywords:
                self.logger.info("未找到关键词配置，检查是否有默认关键词设置")
                
                # 检查是否有硬编码的关键词
                default_import_keywords = [
                    'Calcium Nitrate',
                    'Calcium Nitrate Tetrahydrate', 
                    'Magnesium Nitrate Hexahydrate'
                ]
                
                # 检查这些关键词是否在文本中
                all_config_text = str(self.config.sections())
                for keyword in default_import_keywords:
                    if keyword in all_config_text:
                        import_keywords.append(keyword)
                        self.logger.info(f"从配置中提取到进口关键词: {keyword}")
                
            # 排序和去重
            import_keywords = sorted(list(set(import_keywords)))
            export_keywords = sorted(list(set(export_keywords)))
            
            self.logger.info(f"获取到进口关键词: {import_keywords}")
            self.logger.info(f"获取到出口关键词: {export_keywords}")
            
        except Exception as e:
            self.logger.error(f"获取关键词配置失败: {e}")
            # 返回默认值
            import_keywords = ['Calcium Nitrate', 'Calcium Nitrate Tetrahydrate', 'Magnesium Nitrate Hexahydrate']
            export_keywords = ['Calcium Nitrate', 'Calcium Nitrate Tetrahydrate', 'Magnesium Nitrate Hexahydrate']
        
        return {
            'import': import_keywords,
            'export': export_keywords
        }
    
    def set_keywords(self, import_keywords, export_keywords):
        """保存关键词配置"""
        try:
            # 先清除现有关键词
            for key in list(self.config['keywords'].keys()):
                if key.startswith(('进口关键词', '出口关键词')):
                    del self.config['keywords'][key]
            
            # 添加新关键词
            for i, kw in enumerate(import_keywords, 1):
                self.config.set('keywords', f'进口关键词{i}', kw)
            
            for i, kw in enumerate(export_keywords, 1):
                self.config.set('keywords', f'出口关键词{i}', kw)

            # 自动维护“关键词 -> 中文货名”映射：
            # - 用户只配置关键词
            # - 系统自动补齐未存在的中文映射
            # - 未知关键词默认回填英文关键词（避免Excel中文列为空）
            self._ensure_keyword_translations(list(set(import_keywords + export_keywords)))
            
            # 保存配置
            if self.save_config():
                self.update_runtime_config()
                self.logger.info("关键词配置已保存")
                return True
            else:
                self.logger.error("保存关键词配置失败")
                return False
                
        except Exception as e:
            self.logger.error(f"保存关键词配置失败: {e}")
            return False

    def _auto_translate_keyword(self, keyword: str) -> str:
        """关键词自动中文化（无外部依赖的保底方案）

        说明：
        - 优先使用内置常见危险品关键词映射
        - 若无法识别，则回填英文（避免Excel中文列为空）
        """
        builtin = {
            'Calcium Nitrate': '硝酸钙',
            'Calcium Nitrate Tetrahydrate': '四水合硝酸钙',
            'Magnesium Nitrate Hexahydrate': '六水合硝酸镁'
        }
        return builtin.get(keyword, keyword)

    def _ensure_keyword_translations(self, keywords_list):
        """确保 keyword_translation 节中包含所有关键词的中文映射"""
        try:
            if not self.config.has_section('keyword_translation'):
                self.config.add_section('keyword_translation')

            for kw in keywords_list:
                kw = (kw or '').strip()
                if not kw:
                    continue
                # 若不存在映射，则自动补齐
                if not self.config.has_option('keyword_translation', kw):
                    self.config.set('keyword_translation', kw, self._auto_translate_keyword(kw))
        except Exception as e:
            self.logger.error(f"自动维护关键词中文映射失败: {e}")

    def get_keyword_translation_map(self):
        """获取关键词中文映射字典"""
        try:
            if not self.config.has_section('keyword_translation'):
                return {}
            # ConfigParser 默认会把 key 转小写，需保留原样。
            # 这里通过 items() 读取原始 key（在 ini 中写入的 key 本身）。
            return {k: v for k, v in self.config.items('keyword_translation')}
        except Exception as e:
            self.logger.error(f"获取关键词中文映射失败: {e}")
            return {}

    # （已移除重复的 _sync_keyword_translations / get_keyword_translation_map 实现）
    
    # 短信配置
    def get_sms_config(self):
        """获取短信配置"""
        try:
            return {
                'account': self.config.get('sms', '短信账户', fallback=''),
                'password': self.config.get('sms', '短信密码', fallback=''),
                'mobiles': self.config.get('sms', '短信手机号', fallback=''),
                'import_template': self.config.get('sms', '进口短信模板', fallback=''),
                'export_template': self.config.get('sms', '出口短信模板', fallback=''),
                'api_url': self.config.get('sms', '短信API地址', fallback='')
            }
        except Exception as e:
            self.logger.error(f"获取短信配置失败: {e}")
            return {}
    
    def set_sms_config(self, sms_config):
        """保存短信配置"""
        try:
            for key, value in sms_config.items():
                if key == 'account':
                    self.config.set('sms', '短信账户', str(value))
                elif key == 'password':
                    self.config.set('sms', '短信密码', str(value))
                elif key == 'mobiles':
                    self.config.set('sms', '短信手机号', str(value))
                elif key == 'import_template':
                    self.config.set('sms', '进口短信模板', str(value))
                elif key == 'export_template':
                    self.config.set('sms', '出口短信模板', str(value))
                elif key == 'api_url':
                    self.config.set('sms', '短信API地址', str(value))
            
            # 保存配置
            if self.save_config():
                self.update_runtime_config()
                self.logger.info("短信配置已保存")
                return True
            else:
                self.logger.error("保存短信配置失败")
                return False
                
        except Exception as e:
            self.logger.error(f"保存短信配置失败: {e}")
            return False
    
    # 文件路径配置
    def get_file_paths(self):
        """获取文件路径配置"""
        try:
            return {
                'import_db': self.config.get('files', '进口数据库', fallback='processed_emails_import.db'),
                'export_db': self.config.get('files', '出口数据库', fallback='processed_emails.db'),
                'import_log': self.config.get('files', '进口日志文件', fallback='email_processing_log_import.csv'),
                'export_log': self.config.get('files', '出口日志文件', fallback='email_processing_log.csv')
            }
        except Exception as e:
            self.logger.error(f"获取文件路径配置失败: {e}")
            return {}
    
    # 系统设置
    def get_system_settings(self):
        """获取系统设置"""
        try:
            return {
                'check_interval': self.config.getint('settings', '检查间隔', fallback=30),
                'log_retention_days': self.config.getint('settings', '日志保留天数', fallback=30),
                'db_retention_days': self.config.getint('settings', '数据库保留天数', fallback=90),
                'theme': self.config.get('settings', '界面主题', fallback='dark-blue'),
                'font_size': self.config.getint('settings', '字体大小', fallback=12)
            }
        except Exception as e:
            self.logger.error(f"获取系统设置失败: {e}")
            return {}
    
    def set_system_settings(self, settings):
        """保存系统设置"""
        try:
            for key, value in settings.items():
                if key == 'check_interval':
                    self.config.set('settings', '检查间隔', str(value))
                elif key == 'log_retention_days':
                    self.config.set('settings', '日志保留天数', str(value))
                elif key == 'db_retention_days':
                    self.config.set('settings', '数据库保留天数', str(value))
                elif key == 'theme':
                    self.config.set('settings', '界面主题', value)
                elif key == 'font_size':
                    self.config.set('settings', '字体大小', str(value))
            
            # 保存配置
            if self.save_config():
                self.update_runtime_config()
                self.logger.info("系统设置已保存")
                return True
            else:
                self.logger.error("保存系统设置失败")
                return False
                
        except Exception as e:
            self.logger.error(f"保存系统设置失败: {e}")
            return False
    
    # Web配置
    def get_web_config(self):
        """获取Web配置"""
        try:
            return {
                'host': self.config.get('web', '主机', fallback='0.0.0.0'),
                'port': self.config.getint('web', '端口', fallback=5000),
                'debug': self.config.getboolean('web', '调试模式', fallback=False),
                'stats_days': self.config.getint('web', '统计天数', fallback=30),
                'chart_dpi': self.config.getint('web', '图表DPI', fallback=100),
                'monitor_interval': self.config.getint('web', '监控间隔', fallback=30)
            }
        except Exception as e:
            self.logger.error(f"获取Web配置失败: {e}")
            return {}
    
    # 额外收件人配置
    def get_additional_recipients(self, type_='both'):
        """获取额外收件人列表"""
        recipients = {
            'import': [],
            'export': []
        }
        
        try:
            for key in self.config['additional_recipients']:
                if key.startswith('import_'):
                    recipients['import'].append(self.config['additional_recipients'][key])
                elif key.startswith('export_'):
                    recipients['export'].append(self.config['additional_recipients'][key])
        except Exception as e:
            self.logger.error(f"获取额外收件人配置失败: {e}")
        
        if type_ == 'import':
            return recipients['import']
        elif type_ == 'export':
            return recipients['export']
        return recipients
    
    def set_additional_recipients(self, import_recipients, export_recipients):
        """保存额外收件人配置"""
        try:
            # 先清除现有收件人
            for key in list(self.config['additional_recipients'].keys()):
                if key.startswith(('import_', 'export_')):
                    del self.config['additional_recipients'][key]
            
            # 添加新收件人
            for i, recipient in enumerate(import_recipients, 1):
                self.config.set('additional_recipients', f'import_{i}', recipient)
            
            for i, recipient in enumerate(export_recipients, 1):
                self.config.set('additional_recipients', f'export_{i}', recipient)
            
            # 保存配置
            if self.save_config():
                self.update_runtime_config()
                self.logger.info("额外收件人配置已保存")
                return True
            else:
                self.logger.error("保存额外收件人配置失败")
                return False
                
        except Exception as e:
            self.logger.error(f"保存额外收件人配置失败: {e}")
            return False
    
    # 获取所有配置
    def get_all_configs(self):
        """获取所有配置"""
        try:
            return {
                'email': self.get_email_config(),
                'keywords': self.get_keywords(),
                'sms': self.get_sms_config(),
                'files': self.get_file_paths(),
                'settings': self.get_system_settings(),
                'web': self.get_web_config(),
                'additional_recipients': self.get_additional_recipients()
            }
        except Exception as e:
            self.logger.error(f"获取所有配置失败: {e}")
            return {}