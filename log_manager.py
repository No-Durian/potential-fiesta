# log_manager.py
"""
统一的日志管理器
"""

import os
import logging
import sys
from datetime import datetime
import traceback

class LogManager:
    def __init__(self, app_name="舱单邮件系统"):
        self.app_name = app_name
        self.log_dir = "logs"
        self.setup_logging()
    
    def setup_logging(self):
        """设置日志系统"""
        # 创建日志目录
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir, exist_ok=True)
        
        # 设置日志文件名（每天一个文件）
        today = datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(self.log_dir, f"{today}.log")
        
        # 配置根日志记录器
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        
        # 清除现有的处理器
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        # 创建格式器
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # 文件处理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        
        # 控制台处理器
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        console_handler.setLevel(logging.INFO)
        
        # 添加处理器
        root_logger.addHandler(file_handler)
        root_logger.addHandler(console_handler)
        
        # 设置特定模块的日志级别
        logging.getLogger('urllib3').setLevel(logging.WARNING)
        logging.getLogger('smtplib').setLevel(logging.WARNING)
        
        self.logger = logging.getLogger(self.app_name)
        self.logger.info(f"日志系统初始化完成，日志文件: {log_file}")
    
    def get_logger(self, name=None):
        """获取指定名称的日志记录器"""
        if name:
            return logging.getLogger(f"{self.app_name}.{name}")
        return self.logger
    
    def log_exception(self, context=""):
        """记录异常信息"""
        exc_info = sys.exc_info()
        if exc_info[0] is not None:
            self.logger.error(f"{context}异常: {exc_info[1]}", exc_info=exc_info)
    
    def get_recent_logs(self, lines=100):
        """获取最近的日志内容"""
        today = datetime.now().strftime("%Y%m%d")
        log_file = os.path.join(self.log_dir, f"{today}.log")
        
        if os.path.exists(log_file):
            try:
                with open(log_file, 'r', encoding='utf-8') as f:
                    all_lines = f.readlines()
                    return "".join(all_lines[-lines:])
            except Exception as e:
                return f"读取日志失败: {e}"
        return "日志文件不存在"