from config_manager import ConfigManager

# 初始化配置管理器
config_manager = ConfigManager()
config = config_manager.get_all_configs()

# Web服务器配置
WEB_HOST = config['web']['host']
WEB_PORT = config['web']['port']
DEBUG_MODE = config['web']['debug']

# 数据库文件路径
IMPORT_DB_FILE = config['files']['import_db']
EXPORT_DB_FILE = config['files']['export_db']

# 日志文件路径
IMPORT_LOG_FILE = config['files']['import_log']
EXPORT_LOG_FILE = config['files']['export_log']

# 统计配置
STATS_DAYS = config['web']['stats_days']
CHART_DPI = config['web']['chart_dpi']

# 系统监控间隔（秒）
MONITOR_INTERVAL = config['web']['monitor_interval']