#!/usr/bin/env python3
"""
舱单自动处理系统 - Web管理界面
提供数据库查看、日志管理、短信测试和统计功能
"""

import sys
import os
import sqlite3
import csv
import json
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify, send_file, Response
from flask_cors import CORS
import threading
import time
import logging
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
import matplotlib
matplotlib.use('Agg')  # 使用非GUI后端
import io
import base64
import subprocess
from config_manager import ConfigManager
config_manager = ConfigManager()
import subprocess
import threading

# 全局变量
system_process = None
system_running = False
system_logs = []

# 导入现有的统计系统
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
try:
    from statistics_system import StatisticsSystem
    stats_system = StatisticsSystem()
except ImportError as e:
    print(f"⚠️ 警告: 无法导入统计系统: {e}")
    # 创建一个模拟的统计系统
    class MockStatisticsSystem:
        def __init__(self):
            self.import_db_file = 'processed_emails_import.db'
            self.export_db_file = 'processed_emails.db'
        
        def init_database(self):
            return True
            
        def get_keywords_summary(self, start_date=None, end_date=None):
            return {
                'import': {},
                'export': {},
                'total_import': 0,
                'total_export': 0,
                'all_keywords': []
            }
    
    stats_system = MockStatisticsSystem()

def update_runtime_config(self):
    """更新运行时配置"""
    try:
        # 重新加载配置到各处理模块
        import importlib
        
        # 尝试更新进口模块配置
        try:
            import InputAutoRW_FullFunc_2_0
            importlib.reload(InputAutoRW_FullFunc_2_0)
            logger.info("✅ 进口模块配置已更新")
        except Exception as e:
            logger.warning(f"⚠️ 更新进口模块配置失败: {e}")
        
        # 尝试更新出口模块配置
        try:
            import OutputAutoRWwithSend_3_0
            importlib.reload(OutputAutoRWwithSend_3_0)
            logger.info("✅ 出口模块配置已更新")
        except Exception as e:
            logger.warning(f"⚠️ 更新出口模块配置失败: {e}")
            
        return True
    except Exception as e:
        logger.error(f"❌ 更新运行时配置失败: {e}")
        return False
# 历史邮件同步管理器
class HistoryMailSyncManager:
    """历史邮件同步管理器"""
    def __init__(self):
        pass
    
    def sync_history_emails(self):
        """同步历史邮件"""
        try:
            # 使用UTF-8编码，避免GBK解码错误
            result = subprocess.run(
                [sys.executable, 'HistoryMailSync.py'], 
                capture_output=True, 
                text=True,
                encoding='utf-8',
                errors='ignore',  # 忽略无法解码的字符
                timeout=300  # 5分钟超时
            )
            return {
                'success': True,
                'message': '历史邮件同步完成',
                'output': result.stdout
            }
        except subprocess.TimeoutExpired:
            return {
                'success': False,
                'message': '同步超时，可能需要手动运行 HistoryMailSync.py'
            }
        except FileNotFoundError:
            return {
                'success': False,
                'message': '找不到 HistoryMailSync.py 文件'
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'同步失败: {str(e)}'
            }
    
    def check_database_integrity(self):
        """检查数据库完整性"""
        try:
            # 检查两个数据库的连接和表结构
            issues = []
            
            # 检查进口数据库
            import_db = 'processed_emails_import.db'
            if os.path.exists(import_db):
                conn = sqlite3.connect(import_db)
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails';")
                table_exists = cursor.fetchone()
                if not table_exists:
                    issues.append('进口数据库缺少 keyword_emails 表')
                conn.close()
            else:
                issues.append('进口数据库文件不存在')
            
            # 检查出口数据库
            export_db = 'processed_emails.db'
            if os.path.exists(export_db):
                conn = sqlite3.connect(export_db)
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails';")
                table_exists = cursor.fetchone()
                if not table_exists:
                    issues.append('出口数据库缺少 keyword_emails 表')
                conn.close()
            else:
                issues.append('出口数据库文件不存在')
            
            return {
                'success': True,
                'message': '数据库完整性检查完成',
                'issues': issues if issues else ['所有数据库结构正常']
            }
        except Exception as e:
            return {
                'success': False,
                'message': f'检查失败: {str(e)}'
            }

# 创建历史邮件同步管理器实例
history_sync_manager = HistoryMailSyncManager()

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('web_interface.log', encoding='utf-8')
    ]
)

logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)  # 允许跨域请求

# 系统状态
system_status = {
    'import_running': False,
    'export_running': False,
    'last_check': None
}

def ensure_database_exists():
    """确保数据库表存在"""
    try:
        # 进口数据库
        import_db_file = 'processed_emails_import.db'
        if os.path.exists(import_db_file):
            conn = sqlite3.connect(import_db_file)
            cursor = conn.cursor()
            
            # 检查表是否存在 - 更详细的日志
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails';")
            table_exists = cursor.fetchone()
            
            if not table_exists:
                # 创建与 InputAutoRW_FullFunc_2_0.py 一致的表结构
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
                        english_goods_descriptions TEXT,
                        chinese_goods_descriptions TEXT,
                        sync_source TEXT DEFAULT '',
                        UNIQUE(email_uid)
                    )
                ''')
                # 创建索引
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_email_uid ON keyword_emails(email_uid)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_processed_date ON keyword_emails(processed_date)')
                conn.commit()
                logger.info("创建进口数据库表成功")
            else:
                logger.info("进口数据库表已存在")
                
            conn.close()
        
        # 出口数据库（保持不变）
        export_db_file = 'processed_emails.db'
        if os.path.exists(export_db_file):
            conn = sqlite3.connect(export_db_file)
            cursor = conn.cursor()
            
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='keyword_emails';")
            if not cursor.fetchone():
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS keyword_emails (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        email_uid TEXT NOT NULL,
                        sender TEXT,
                        sender_address TEXT,
                        subject TEXT,
                        received_date TEXT,
                        processed_date TEXT,
                        matched_keywords TEXT,
                        excel_sent INTEGER DEFAULT 0,
                        txt_attachment TEXT,
                        container_count INTEGER DEFAULT 0,
                        attachment_names TEXT
                    )
                ''')
                conn.commit()
                logger.info("创建出口数据库表")
            
            conn.close()
            
        logger.info("数据库检查完成")
        
    except Exception as e:
        logger.error(f"确保数据库存在失败: {e}", exc_info=True)

class SystemMonitor(threading.Thread):
    """系统状态监控线程"""
    def __init__(self):
        super().__init__()
        self.daemon = True
        self.running = True
        
    def run(self):
        while self.running:
            try:
                # 检查处理程序是否在运行
                import_running = False
                export_running = False
                
                # 检查进口处理程序
                try:
                    # 尝试导入模块，但不真正调用函数
                    import InputAutoRW_FullFunc_2_0
                    import_running = True
                except:
                    import_running = False
                
                # 检查出口处理程序
                try:
                    import OutputAutoRWwithSend_3_0
                    export_running = True
                except:
                    export_running = False
                
                system_status['import_running'] = import_running
                system_status['export_running'] = export_running
                system_status['last_check'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
            except Exception as e:
                logger.error(f"监控系统状态失败: {e}")
            
            time.sleep(30)  # 30秒检查一次

# 启动监控线程
monitor = SystemMonitor()
monitor.start()



@app.route('/api/config', methods=['GET'])
def get_config():
    """获取所有配置"""
    try:
        config = config_manager.get_all_configs()
        return jsonify({
            'success': True,
            'data': config
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        })

@app.route('/api/config/keywords', methods=['POST'])
def save_keywords():
    """保存关键词配置"""
    if check_system_running():
        return jsonify({
            'success': False,
            'message': '请先停止系统再修改配置'
        })
    
    try:
        # 确保请求数据是 JSON 格式
        if not request.is_json:
            return jsonify({
                'success': False,
                'message': '请求必须是 JSON 格式'
            }), 415
        
        data = request.get_json()
        import_keywords = data.get('import', [])
        export_keywords = data.get('export', [])
        
        if config_manager.set_keywords(import_keywords, export_keywords):
            return jsonify({
                'success': True,
                'message': '关键词配置已保存'
            })
        else:
            return jsonify({
                'success': False,
                'message': '保存关键词配置失败'
            })
    except Exception as e:
        logger.error(f"保存关键词配置失败: {e}")
        return jsonify({
            'success': False,
            'message': str(e)
        })

@app.route('/api/config/sms', methods=['POST'])
def save_sms_config():
    """保存短信配置"""
    if check_system_running():
        return jsonify({
            'success': False,
            'message': '请先停止系统再修改配置'
        })
    
    try:
        # 确保请求数据是 JSON 格式
        if not request.is_json:
            return jsonify({
                'success': False,
                'message': '请求必须是 JSON 格式'
            }), 415
        
        sms_config = request.get_json()
        
        if config_manager.set_sms_config(sms_config):
            return jsonify({
                'success': True,
                'message': '短信配置已保存'
            })
        else:
            return jsonify({
                'success': False,
                'message': '保存短信配置失败'
            })
    except Exception as e:
        logger.error(f"保存短信配置失败: {e}")
        return jsonify({
            'success': False,
            'message': str(e)
        })

@app.route('/api/config/settings', methods=['POST'])
def save_system_settings():
    """保存系统设置"""
    if check_system_running():
        return jsonify({
            'success': False,
            'message': '请先停止系统再修改配置'
        })
    
    try:
        # 确保请求数据是 JSON 格式
        if not request.is_json:
            return jsonify({
                'success': False,
                'message': '请求必须是 JSON 格式'
            }), 415
        
        settings = request.get_json()
        
        if config_manager.set_system_settings(settings):
            return jsonify({
                'success': True,
                'message': '系统设置已保存'
            })
        else:
            return jsonify({
                'success': False,
                'message': '保存系统设置失败'
            })
    except Exception as e:
        logger.error(f"保存系统设置失败: {e}")
        return jsonify({
            'success': False,
            'message': str(e)
        })
    


@app.route('/config')
def config_page():
    return render_template('config_fixed.html')  # 使用修复后的配置页面



@app.route('/api/config/recipients', methods=['POST'])
def save_recipients():
    """保存额外收件人配置"""
    if check_system_running():
        return jsonify({
            'success': False,
            'message': '请先停止系统再修改配置'
        })
    
    try:
        # 确保请求数据是 JSON 格式
        if not request.is_json:
            return jsonify({
                'success': False,
                'message': '请求必须是 JSON 格式'
            }), 415
        
        data = request.get_json()
        import_recipients = data.get('import', [])
        export_recipients = data.get('export', [])
        
        if config_manager.set_additional_recipients(import_recipients, export_recipients):
            return jsonify({
                'success': True,
                'message': '额外收件人配置已保存'
            })
        else:
            return jsonify({
                'success': False,
                'message': '保存额外收件人配置失败'
            })
    except Exception as e:
        logger.error(f"保存额外收件人配置失败: {e}")
        return jsonify({
            'success': False,
            'message': str(e)
        })










@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

# 修改原有的系统状态接口
@app.route('/api/system/status')
def get_system_status():
    """获取系统状态"""
    try:
        # 检查主控制器是否在运行
        import_running = False
        export_running = False
        
        # 检查进程是否在运行
        if system_process and system_process.poll() is None:
            import_running = True
            export_running = True
        else:
            # 检查线程是否在运行
            try:
                for thread in threading.enumerate():
                    if 'ImportProcessor' in thread.name:
                        import_running = True
                    if 'ExportProcessor' in thread.name:
                        export_running = True
            except:
                pass
        
        system_status['import_running'] = import_running
        system_status['export_running'] = export_running
        system_status['last_check'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        import_db_file = 'processed_emails_import.db'
        export_db_file = 'processed_emails.db'
        
        import_count = 0
        export_count = 0
        today_import = 0
        today_export = 0
        
        # 进口数据库统计
        if os.path.exists(import_db_file):
            try:
                conn = sqlite3.connect(import_db_file)
                cursor = conn.cursor()
                cursor.execute('SELECT COUNT(*) FROM keyword_emails')
                import_count = cursor.fetchone()[0]
                
                # 获取今日统计
                today = datetime.now().strftime('%Y-%m-%d')
                cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE DATE(processed_date) = ?', (today,))
                today_import = cursor.fetchone()[0]
                
                conn.close()
            except Exception as e:
                logger.warning(f"进口数据库查询失败: {e}")
        
        # 出口数据库统计
        if os.path.exists(export_db_file):
            try:
                conn = sqlite3.connect(export_db_file)
                cursor = conn.cursor()
                cursor.execute('SELECT COUNT(*) FROM keyword_emails')
                export_count = cursor.fetchone()[0]
                
                # 获取今日统计
                today = datetime.now().strftime('%Y-%m-%d')
                cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE DATE(processed_date) = ?', (today,))
                today_export = cursor.fetchone()[0]
                
                conn.close()
            except Exception as e:
                logger.warning(f"出口数据库查询失败: {e}")
        
        status = {
            'system': {
                'import_running': system_status['import_running'],
                'export_running': system_status['export_running'],
                'last_check': system_status['last_check']
            },
            'database': {
                'import_total': import_count,
                'export_total': export_count,
                'today_import': today_import,
                'today_export': today_export
            }
        }
        
        return jsonify({'success': True, 'data': status})
    except Exception as e:
        logger.error(f"获取系统状态失败: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)})

#新增：检查系统是否运行20251222
def check_system_running():
    """检查系统是否在运行"""
    if system_running:
        return True
    return False

@app.route('/api/config/email', methods=['POST'])
def save_email_config():
    """保存邮箱配置"""
    if check_system_running():
        return jsonify({
            'success': False,
            'message': '请先停止系统再修改配置'
        })
    
    try:
        # 确保请求数据是 JSON 格式
        if not request.is_json:
            return jsonify({
                'success': False,
                'message': '请求必须是 JSON 格式'
            }), 415
        
        email_config = request.get_json()
        
        # 验证必要字段
        required_fields = ['import_email', 'import_password', 'export_email', 'export_password']
        for field in required_fields:
            if field not in email_config or not email_config[field]:
                return jsonify({
                    'success': False,
                    'message': f'缺少必要字段: {field}'
                })
        
        if config_manager.set_email_config(email_config):
            return jsonify({
                'success': True,
                'message': '邮箱配置已保存'
            })
        else:
            return jsonify({
                'success': False,
                'message': '保存邮箱配置失败'
            })
    except Exception as e:
        logger.error(f"保存邮箱配置失败: {e}")
        return jsonify({
            'success': False,
            'message': str(e)
        })






















@app.route('/api/system/sync_history')
def sync_history():
    """同步历史邮件"""
    try:
        result = history_sync_manager.sync_history_emails()
        return jsonify(result)
    except Exception as e:
        logger.error(f"同步历史邮件失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/system/check_database')
def check_database():
    """检查数据库完整性"""
    try:
        result = history_sync_manager.check_database_integrity()
        return jsonify(result)
    except Exception as e:
        logger.error(f"检查数据库失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/import/database')
def get_import_database():
    """获取进口数据库记录"""
    try:
        logger.info("开始查询进口数据库...")
        
        page = int(request.args.get('page', 1))
        page_size = int(request.args.get('page_size', 50))
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')
        keywords = request.args.get('keywords', '')
        
        offset = (page - 1) * page_size
        
        db_file = 'processed_emails_import.db'
        
        if not os.path.exists(db_file):
            logger.warning(f"进口数据库文件不存在: {db_file}")
            return jsonify({'success': True, 'data': [], 'total': 0, 'page': page})
        
        conn = sqlite3.connect(db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 构建查询条件
        conditions = []
        params = []
        
        if start_date:
            # 使用正确的列名 processed_date
            conditions.append("DATE(processed_date) >= ?")
            params.append(start_date)
        
        if end_date:
            # 使用正确的列名 processed_date
            conditions.append("DATE(processed_date) <= ?")
            params.append(end_date)
        
        if keywords:
            keyword_list = [k.strip() for k in keywords.split(',') if k.strip()]
            if keyword_list:
                keyword_conditions = []
                for keyword in keyword_list:
                    keyword_conditions.append("matched_keywords LIKE ?")
                    params.append(f'%{keyword}%')
                conditions.append(f"({' OR '.join(keyword_conditions)})")
        
        where_clause = " AND ".join(conditions) if conditions else "1=1"
        
        # 获取总数
        try:
            count_query = f"SELECT COUNT(*) FROM keyword_emails WHERE {where_clause}"
            cursor.execute(count_query, params)
            total = cursor.fetchone()[0]
        except sqlite3.OperationalError as e:
            logger.error(f"查询总数失败: {e}")
            total = 0
        
        # 获取分页数据 - 使用正确的列名
        query = f"""
        SELECT id, email_uid, sender, sender_address, subject, received_date,
               processed_date, matched_keywords, excel_sent, txt_attachment,
               container_count, attachment_names, sync_source,
               english_goods_descriptions, chinese_goods_descriptions
        FROM keyword_emails
        WHERE {where_clause}
        ORDER BY processed_date DESC
        LIMIT ? OFFSET ?
        """
        
        params.extend([page_size, offset])
        
        data = []
        try:
            cursor.execute(query, params)
            rows = cursor.fetchall()
            # 转换为字典列表
            data = [dict(row) for row in rows]
            logger.info(f"成功查询到 {len(data)} 条进口记录")
        except sqlite3.OperationalError as e:
            logger.error(f"查询数据失败: {e}")
            data = []
        
        conn.close()
        
        return jsonify({
            'success': True,
            'data': data,
            'total': total,
            'page': page,
            'page_size': page_size,
            'total_pages': (total + page_size - 1) // page_size if total > 0 else 0
        })
    except Exception as e:
        logger.error(f"获取进口数据库失败: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/export/database')
def get_export_database():
    """获取出口数据库记录"""
    try:
        # 确保数据库表存在
        ensure_database_exists()
        
        page = int(request.args.get('page', 1))
        page_size = int(request.args.get('page_size', 50))
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')
        keywords = request.args.get('keywords', '')
        
        offset = (page - 1) * page_size
        
        db_file = 'processed_emails.db'
        
        if not os.path.exists(db_file):
            return jsonify({'success': True, 'data': [], 'total': 0, 'page': page})
        
        conn = sqlite3.connect(db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # 构建查询条件
        conditions = []
        params = []
        
        if start_date:
            conditions.append("DATE(processed_date) >= ?")
            params.append(start_date)
        
        if end_date:
            conditions.append("DATE(processed_date) <= ?")
            params.append(end_date)
        
        if keywords:
            keyword_list = [k.strip() for k in keywords.split(',') if k.strip()]
            if keyword_list:
                keyword_conditions = []
                for keyword in keyword_list:
                    keyword_conditions.append("matched_keywords LIKE ?")
                    params.append(f'%{keyword}%')
                conditions.append(f"({' OR '.join(keyword_conditions)})")
        
        where_clause = " AND ".join(conditions) if conditions else "1=1"
        
        # 获取总数
        try:
            count_query = f"SELECT COUNT(*) FROM keyword_emails WHERE {where_clause}"
            cursor.execute(count_query, params)
            total = cursor.fetchone()[0]
        except sqlite3.OperationalError:
            total = 0
        
        # 获取分页数据
        query = f"""
        SELECT id, email_uid, sender, sender_address, subject, received_date,
               processed_date, matched_keywords, excel_sent, txt_attachment,
               container_count, attachment_names
        FROM keyword_emails
        WHERE {where_clause}
        ORDER BY processed_date DESC
        LIMIT ? OFFSET ?
        """
        
        params.extend([page_size, offset])
        
        data = []
        try:
            cursor.execute(query, params)
            rows = cursor.fetchall()
            # 转换为字典列表
            data = [dict(row) for row in rows]
        except sqlite3.OperationalError:
            data = []
        
        conn.close()
        
        return jsonify({
            'success': True,
            'data': data,
            'total': total,
            'page': page,
            'page_size': page_size,
            'total_pages': (total + page_size - 1) // page_size if total > 0 else 0
        })
    except Exception as e:
        logger.error(f"获取出口数据库失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/statistics/keywords')
def get_keyword_statistics():
    """获取关键词统计"""
    try:
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')
        
        # 手动统计进口数据库
        import_stats = {}
        export_stats = {}
        all_keywords = set()
        
        # 统计进口数据库 - 使用正确的列名
        import_db_file = 'processed_emails_import.db'
        if os.path.exists(import_db_file):
            try:
                conn = sqlite3.connect(import_db_file)
                cursor = conn.cursor()
                
                # 构建查询条件 - 使用 processed_date
                conditions = []
                params = []
                
                if start_date:
                    conditions.append("DATE(processed_date) >= ?")
                    params.append(start_date)
                
                if end_date:
                    conditions.append("DATE(processed_date) <= ?")
                    params.append(end_date)
                
                where_clause = " AND ".join(conditions) if conditions else "1=1"
                
                # 查询匹配关键词
                query = f"""
                SELECT matched_keywords, COUNT(*) as count
                FROM keyword_emails
                WHERE {where_clause}
                GROUP BY matched_keywords
                """
                
                cursor.execute(query, params)
                for keyword_str, count in cursor.fetchall():
                    if keyword_str:
                        # 关键词可能是逗号分隔的多个
                        keywords = [k.strip() for k in keyword_str.split(',') if k.strip()]
                        for keyword in keywords:
                            import_stats[keyword] = import_stats.get(keyword, 0) + 1
                            all_keywords.add(keyword)
                
                conn.close()
            except Exception as e:
                logger.error(f"统计进口关键词失败: {e}")
        
        # 统计出口数据库 - 保持原有逻辑
        export_db_file = 'processed_emails.db'
        if os.path.exists(export_db_file):
            try:
                conn = sqlite3.connect(export_db_file)
                cursor = conn.cursor()
                
                # 构建查询条件
                conditions = []
                params = []
                
                if start_date:
                    conditions.append("DATE(processed_date) >= ?")
                    params.append(start_date)
                
                if end_date:
                    conditions.append("DATE(processed_date) <= ?")
                    params.append(end_date)
                
                where_clause = " AND ".join(conditions) if conditions else "1=1"
                
                # 查询匹配关键词
                query = f"""
                SELECT matched_keywords, COUNT(*) as count
                FROM keyword_emails
                WHERE {where_clause}
                GROUP BY matched_keywords
                """
                
                cursor.execute(query, params)
                for keyword_str, count in cursor.fetchall():
                    if keyword_str:
                        # 关键词可能是逗号分隔的多个
                        keywords = [k.strip() for k in keyword_str.split(',') if k.strip()]
                        for keyword in keywords:
                            export_stats[keyword] = export_stats.get(keyword, 0) + 1
                            all_keywords.add(keyword)
                
                conn.close()
            except Exception as e:
                logger.error(f"统计出口关键词失败: {e}")
        
        # 获取总数 - 进口使用 processed_date
        total_import = 0
        total_export = 0
        
        if os.path.exists(import_db_file):
            try:
                conn = sqlite3.connect(import_db_file)
                cursor = conn.cursor()
                
                conditions = []
                params = []
                
                if start_date:
                    conditions.append("DATE(processed_date) >= ?")
                    params.append(start_date)
                
                if end_date:
                    conditions.append("DATE(processed_date) <= ?")
                    params.append(end_date)
                
                where_clause = " AND ".join(conditions) if conditions else "1=1"
                
                cursor.execute(f"SELECT COUNT(*) FROM keyword_emails WHERE {where_clause}", params)
                total_import = cursor.fetchone()[0] or 0
                conn.close()
            except:
                total_import = 0
        
        if os.path.exists(export_db_file):
            try:
                conn = sqlite3.connect(export_db_file)
                cursor = conn.cursor()
                
                conditions = []
                params = []
                
                if start_date:
                    conditions.append("DATE(processed_date) >= ?")
                    params.append(start_date)
                
                if end_date:
                    conditions.append("DATE(processed_date) <= ?")
                    params.append(end_date)
                
                where_clause = " AND ".join(conditions) if conditions else "1=1"
                
                cursor.execute(f"SELECT COUNT(*) FROM keyword_emails WHERE {where_clause}", params)
                total_export = cursor.fetchone()[0] or 0
                conn.close()
            except:
                total_export = 0
        
        return jsonify({
            'success': True,
            'data': {
                'import': import_stats,
                'export': export_stats,
                'total_import': total_import,
                'total_export': total_export,
                'all_keywords': sorted(list(all_keywords))
            }
        })
    except Exception as e:
        logger.error(f"获取关键词统计失败: {e}", exc_info=True)
        return jsonify({
            'success': True,
            'data': {
                'import': {},
                'export': {},
                'total_import': 0,
                'total_export': 0,
                'all_keywords': []
            }
        })

@app.route('/api/statistics/daily')
def get_daily_statistics():
    """获取每日统计"""
    try:
        days = int(request.args.get('days', 30))
        db_type = request.args.get('type', 'all')
        
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days-1)
        
        dates = []
        import_counts = []
        export_counts = []
        
        for i in range(days):
            current_date = (start_date + timedelta(days=i)).strftime('%Y-%m-%d')
            dates.append(current_date)
            
            # 进口统计 - 使用正确的列名 processed_date
            if db_type in ['all', 'import']:
                import_count = 0
                db_file = 'processed_emails_import.db'
                if os.path.exists(db_file):
                    try:
                        conn = sqlite3.connect(db_file)
                        cursor = conn.cursor()
                        # 使用 DATE 函数处理 processed_date
                        cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE DATE(processed_date) = ?', (current_date,))
                        result = cursor.fetchone()
                        import_count = result[0] if result else 0
                        conn.close()
                    except Exception as e:
                        logger.error(f"查询进口数据库失败 {current_date}: {e}")
                        import_count = 0
                import_counts.append(import_count)
            
            # 出口统计
            if db_type in ['all', 'export']:
                export_count = 0
                db_file = 'processed_emails.db'
                if os.path.exists(db_file):
                    try:
                        conn = sqlite3.connect(db_file)
                        cursor = conn.cursor()
                        cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE DATE(processed_date) = ?', (current_date,))
                        result = cursor.fetchone()
                        export_count = result[0] if result else 0
                        conn.close()
                    except Exception as e:
                        logger.error(f"查询出口数据库失败 {current_date}: {e}")
                        export_count = 0
                export_counts.append(export_count)
        
        data = {
            'dates': dates,
            'import_counts': import_counts if import_counts else None,
            'export_counts': export_counts if export_counts else None
        }
        
        return jsonify({'success': True, 'data': data})
    except Exception as e:
        logger.error(f"获取每日统计失败: {e}", exc_info=True)
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/chart/daily')
def get_daily_chart():
    """生成每日统计图表"""
    try:
        # 确保数据库表存在
        ensure_database_exists()
        
        days = int(request.args.get('days', 30))
        db_type = request.args.get('type', 'all')
        
        # 直接调用函数获取数据，而不是通过 HTTP 请求
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days-1)
        
        dates = []
        import_counts = []
        export_counts = []
        
        for i in range(days):
            current_date = (start_date + timedelta(days=i)).strftime('%Y-%m-%d')
            dates.append(current_date)
            
            # 进口统计
            if db_type in ['all', 'import']:
                import_count = 0
                db_file = 'processed_emails_import.db'
                if os.path.exists(db_file):
                    try:
                        conn = sqlite3.connect(db_file)
                        cursor = conn.cursor()
                        cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE DATE(processed_date) = ?', (current_date,))
                        import_count = cursor.fetchone()[0]
                        conn.close()
                    except:
                        import_count = 0
                import_counts.append(import_count)
            
            # 出口统计
            if db_type in ['all', 'export']:
                export_count = 0
                db_file = 'processed_emails.db'
                if os.path.exists(db_file):
                    try:
                        conn = sqlite3.connect(db_file)
                        cursor = conn.cursor()
                        cursor.execute('SELECT COUNT(*) FROM keyword_emails WHERE DATE(processed_date) = ?', (current_date,))
                        export_count = cursor.fetchone()[0]
                        conn.close()
                    except:
                        export_count = 0
                export_counts.append(export_count)
        
        chart_data = {
            'dates': dates,
            'import_counts': import_counts if import_counts else None,
            'export_counts': export_counts if export_counts else None
        }
        
        # 创建图表
        fig, ax = plt.subplots(figsize=(12, 6))
        
        if db_type in ['all', 'import'] and chart_data['import_counts']:
            ax.plot(chart_data['dates'], chart_data['import_counts'], 
                   marker='o', label='进口', color='#3498db', linewidth=2)
        
        if db_type in ['all', 'export'] and chart_data['export_counts']:
            ax.plot(chart_data['dates'], chart_data['export_counts'], 
                   marker='s', label='出口', color='#2ecc71', linewidth=2)
        
        # 设置图表样式
        ax.set_title(f'最近{days}天舱单处理统计', fontsize=16, fontweight='bold')
        ax.set_xlabel('日期', fontsize=12)
        ax.set_ylabel('处理数量', fontsize=12)
        ax.legend(fontsize=12)
        ax.grid(True, alpha=0.3)
        
        # 旋转x轴标签
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        # 将图表转换为图片
        img = io.BytesIO()
        plt.savefig(img, format='png', dpi=100)
        img.seek(0)
        plt.close(fig)
        
        # 返回Base64编码的图片
        img_base64 = base64.b64encode(img.getvalue()).decode('utf-8')
        
        return jsonify({'success': True, 'image': f'data:image/png;base64,{img_base64}'})
    except Exception as e:
        logger.error(f"生成图表失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/chart/keywords')
def get_keywords_chart():
    """生成关键词统计图表"""
    try:
        # 确保数据库表存在
        ensure_database_exists()
        
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')
        
        # 直接调用函数获取数据，而不是通过 HTTP 请求
        stats_system.init_database()
        result = stats_system.get_keywords_summary(
            start_date if start_date else None,
            end_date if end_date else None
        )
        
        if not result:
            # 如果统计系统失败，使用空数据
            stats = {
                'import': {},
                'export': {},
                'total_import': 0,
                'total_export': 0,
                'all_keywords': []
            }
        else:
            stats = {
                'import': result['import'],
                'export': result['export'],
                'total_import': result['total_import'],
                'total_export': result['total_export'],
                'all_keywords': result['all_keywords']
            }
        
        # 创建图表
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 6))
        
        # 进口关键词统计
        if stats['import']:
            keywords = list(stats['import'].keys())[:10]  # 最多显示10个
            counts = [stats['import'][k] for k in keywords]
            
            ax1.barh(keywords, counts, color='#3498db')
            ax1.set_title('进口关键词统计', fontsize=14, fontweight='bold')
            ax1.set_xlabel('出现次数', fontsize=12)
        else:
            ax1.text(0.5, 0.5, '暂无进口关键词数据', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax1.transAxes, fontsize=12)
            ax1.set_title('进口关键词统计', fontsize=14, fontweight='bold')
        
        # 出口关键词统计
        if stats['export']:
            keywords = list(stats['export'].keys())[:10]  # 最多显示10个
            counts = [stats['export'][k] for k in keywords]
            
            ax2.barh(keywords, counts, color='#2ecc71')
            ax2.set_title('出口关键词统计', fontsize=14, fontweight='bold')
            ax2.set_xlabel('出现次数', fontsize=12)
        else:
            ax2.text(0.5, 0.5, '暂无出口关键词数据', 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax2.transAxes, fontsize=12)
            ax2.set_title('出口关键词统计', fontsize=14, fontweight='bold')
        
        plt.tight_layout()
        
        # 将图表转换为图片
        img = io.BytesIO()
        plt.savefig(img, format='png', dpi=100)
        img.seek(0)
        plt.close(fig)
        
        # 返回Base64编码的图片
        img_base64 = base64.b64encode(img.getvalue()).decode('utf-8')
        
        return jsonify({'success': True, 'image': f'data:image/png;base64,{img_base64}'})
    except Exception as e:
        logger.error(f"生成关键词图表失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/logs/import')
def get_import_logs():
    """获取进口日志"""
    try:
        lines = int(request.args.get('lines', 100))
        log_file = 'email_processing_log_import.csv'
        
        if not os.path.exists(log_file):
            return jsonify({'success': True, 'data': []})
        
        # 读取最后N行
        with open(log_file, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
        
        # 反转顺序，最新的在前
        log_lines = all_lines[-lines:] if lines < len(all_lines) else all_lines
        
        # 解析CSV
        logs = []
        reader = csv.reader(log_lines)
        
        for row in reader:
            if len(row) >= 8:
                logs.append({
                    'timestamp': row[0],
                    'email_uid': row[1],
                    'sender': row[2],
                    'subject': row[3],
                    'has_keyword': row[4] == '1',
                    'excel_sent': row[5] == '1',
                    'matched_keywords': row[6],
                    'container_count': row[7]
                })
        
        return jsonify({'success': True, 'data': logs})
    except Exception as e:
        logger.error(f"获取进口日志失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/logs/export')
def get_export_logs():
    """获取出口日志"""
    try:
        lines = int(request.args.get('lines', 100))
        log_file = 'email_processing_log.csv'
        
        if not os.path.exists(log_file):
            return jsonify({'success': True, 'data': []})
        
        # 读取最后N行
        with open(log_file, 'r', encoding='utf-8') as f:
            all_lines = f.readlines()
        
        # 反转顺序，最新的在前
        log_lines = all_lines[-lines:] if lines < len(all_lines) else all_lines
        
        # 解析CSV
        logs = []
        reader = csv.reader(log_lines)
        
        for row in reader:
            if len(row) >= 8:
                logs.append({
                    'timestamp': row[0],
                    'email_uid': row[1],
                    'sender': row[2],
                    'subject': row[3],
                    'has_keyword': row[4] == '1',
                    'excel_sent': row[5] == '1',
                    'matched_keywords': row[6],
                    'container_count': row[7]
                })
        
        return jsonify({'success': True, 'data': logs})
    except Exception as e:
        logger.error(f"获取出口日志失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/sms/test')
def test_sms():
    """测试短信发送"""
    try:
        sms_type = request.args.get('type', 'import')  # import, export
        
        if sms_type == 'import':
            try:
                from InputAutoRW_FullFunc_2_0 import send_exit_notification
            except ImportError:
                return jsonify({'success': False, 'error': '无法导入进口短信模块'})
        else:
            try:
                from OutputAutoRWwithSend_3_0 import send_exit_notification
            except ImportError:
                return jsonify({'success': False, 'error': '无法导入出口短信模块'})
        
        # 发送测试短信
        result = send_exit_notification("测试短信", is_manual=False)
        
        return jsonify({
            'success': result,
            'message': '短信发送成功' if result else '短信发送失败'
        })
    except Exception as e:
        logger.error(f"测试短信发送失败: {e}")
        return jsonify({'success': False, 'error': str(e)})




# 添加新的配置页面路由
@app.route('/config/fixed')
def config_fixed_page():
    return render_template('config_fixed.html')

# 添加配置检查路由
@app.route('/api/config/check')
def check_config_status():
    """检查配置状态"""
    try:
        # 检查系统是否运行
        system_running = check_system_running()
        
        # 检查配置文件是否存在
        config_exists = os.path.exists('config.ini')
        
        return jsonify({
            'success': True,
            'data': {
                'system_running': system_running,
                'config_exists': config_exists,
                'can_modify': not system_running
            }
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'message': str(e)
        })








@app.route('/api/export/csv')
def export_csv():
    """导出CSV文件"""
    try:
        # 确保数据库表存在
        ensure_database_exists()
        
        db_type = request.args.get('type', 'import')
        start_date = request.args.get('start_date', '')
        end_date = request.args.get('end_date', '')
        
        if db_type == 'import':
            db_file = 'processed_emails_import.db'
            filename = f'进口舱单统计_{start_date}_至_{end_date}.csv'
        else:
            db_file = 'processed_emails.db'
            filename = f'出口舱单统计_{start_date}_至_{end_date}.csv'
        
        if not os.path.exists(db_file):
            return jsonify({'success': False, 'error': '数据库文件不存在'})
        
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()
        
        # 构建查询条件
        conditions = []
        params = []
        
        if start_date:
            conditions.append("DATE(processed_date) >= ?")
            params.append(start_date)
        
        if end_date:
            conditions.append("DATE(processed_date) <= ?")
            params.append(end_date)
        
        where_clause = " AND ".join(conditions) if conditions else "1=1"
        
        # 查询数据
        query = f"""
        SELECT processed_date, sender, subject, matched_keywords, 
               container_count, attachment_names
        FROM keyword_emails
        WHERE {where_clause}
        ORDER BY processed_date DESC
        """
        
        try:
            cursor.execute(query, params)
            rows = cursor.fetchall()
        except sqlite3.OperationalError:
            rows = []
        
        # 创建CSV文件
        output = io.StringIO()
        writer = csv.writer(output)
        
        # 写入标题行
        writer.writerow(['处理日期', '发件人', '邮件主题', '匹配关键词', '箱数', '附件名称'])
        
        # 写入数据
        for row in rows:
            writer.writerow(row)
        
        conn.close()
        
        # 返回CSV文件
        output.seek(0)
        return Response(
            output.getvalue(),
            mimetype="text/csv",
            headers={"Content-disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        logger.error(f"导出CSV失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/api/settings/keywords')
def get_keywords():
    """获取关键词设置"""
    try:
        # 直接从配置文件读取（避免模块级变量缓存导致的“配置更新但页面不更新”问题）
        kw_cfg = config_manager.get_keywords()
        import_keywords = kw_cfg.get('import', [])
        export_keywords = kw_cfg.get('export', [])
        
        return jsonify({
            'success': True,
            'data': {
                'import': import_keywords,
                'export': export_keywords
            }
        })
    except Exception as e:
        logger.error(f"获取关键词失败: {e}")
        return jsonify({'success': False, 'error': str(e)})

# 修改启动函数，确保不会自动启动
@app.route('/api/system/start')
def start_system():
    """启动处理系统"""
    global system_process, system_running
    
    try:
        # 确保数据库表存在
        ensure_database_exists()
        
        if system_running:
            return jsonify({'success': False, 'error': '系统已经在运行'})
        
        logger.info("🚀 用户手动启动系统...")

        # ===== 启动前预检：数据库结构升级 + 历史邮件同步 =====
        # 目的：把“已手动发送过的舱单邮件”同步进数据库，避免自动回复时重复发送。
        try:
            # 1) 数据库结构升级（补齐 sync_source 列）
            upgrade = subprocess.run(
                [sys.executable, 'UpdateDatabaseSchema.py'],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore',
                timeout=120
            )
            if upgrade.stdout:
                system_logs.append("[预检] 数据库结构升级完成")
            if upgrade.returncode != 0:
                logger.warning(f"⚠️ 数据库升级脚本返回非0: {upgrade.returncode} | {upgrade.stderr}")

            # 2) 同步历史邮件（必须步骤，失败则阻止启动，避免重复发送）
            from HistoryMailSync import HistoryMailSync
            sync_mgr = HistoryMailSync()
            sync_res = sync_mgr.sync_all_folders(max_emails=100, progress_callback=None)
            if not sync_res.get('success'):
                return jsonify({'success': False, 'error': f"启动前历史邮件同步失败：{sync_res.get('message','未知错误')}"})

            # 把同步输出写入系统日志（最多保留最后200行，避免页面太长）
            out = (sync_res.get('output') or '').strip()
            if out:
                tail_lines = out.splitlines()[-200:]
                system_logs.extend([f"[预检] {l}" for l in tail_lines if l.strip()])
                if len(system_logs) > 1000:
                    system_logs[:] = system_logs[-1000:]
        except Exception as e:
            return jsonify({'success': False, 'error': f"启动前预检失败：{e}"})
        
        # 启动主控制器
        system_process = subprocess.Popen(
            [sys.executable, 'AutoRW_MainController_fixed.py', 'start'],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            bufsize=1,
            universal_newlines=True
        )
        
        system_running = True
        
        # 启动线程读取输出
        def read_output():
            global system_logs
            while system_running and system_process.poll() is None:
                try:
                    line = system_process.stdout.readline()
                    if line:
                        clean_line = line.strip()
                        system_logs.append(clean_line)
                        logger.info(f"[系统] {clean_line}")
                        
                        # 保持日志最多1000行
                        if len(system_logs) > 1000:
                            system_logs = system_logs[-1000:]
                except Exception as e:
                    logger.error(f"读取系统输出失败: {e}")
                    break
        
        output_thread = threading.Thread(target=read_output, daemon=True)
        output_thread.start()
        
        # 等待一会儿确认启动
        time.sleep(2)
        
        return jsonify({
            'success': True,
            'message': '系统已启动',
            'pid': system_process.pid
        })

    except Exception as e:
        logger.error(f"启动系统失败: {e}")
        system_running = False
        return jsonify({'success': False, 'error': str(e)})


#新增：停止程序20251222
@app.route('/api/system/stop')
def stop_system():
    """停止系统"""
    global system_process, system_running
    
    try:
        if not system_running or not system_process:
            return jsonify({'success': False, 'error': '系统未运行'})
        
        # 发送停止信号
        if system_process and system_process.poll() is None:
            system_process.terminate()
            try:
                system_process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                system_process.kill()
                system_process.wait()
        
        system_running = False
        system_process = None
        
        return jsonify({
            'success': True,
            'message': '系统已停止'
        })
        
    except Exception as e:
        logger.error(f"停止系统失败: {e}")
        return jsonify({'success': False, 'error': str(e)})


#新增：获取系统日志20251222
@app.route('/api/system/logs')
def get_system_logs():
    """获取系统日志"""
    global system_logs
    
    try:
        # 如果是运行中，也检查是否有新日志
        if system_process and system_process.poll() is None:
            try:
                # 尝试读取最新输出
                while True:
                    line = system_process.stdout.readline()
                    if line:
                        system_logs.append(line.strip())
                    else:
                        break
            except:
                pass
        
        return jsonify({
            'success': True,
            'logs': '\n'.join(system_logs[-200:])  # 返回最近200行
        })
        
    except Exception as e:
        logger.error(f"获取系统日志失败: {e}")
        return jsonify({'success': True, 'logs': '无法获取系统日志'})

@app.route('/api/system/logs/clear')
def clear_system_logs():
    """清空系统日志"""
    global system_logs
    system_logs = []
    return jsonify({'success': True, 'message': '日志已清空'})














def run_import_processor():
    """运行进口处理程序"""
    try:
        # 使用主控制器启动进口处理程序
        subprocess.run([sys.executable, 'InputAutoRW_FullFunc_2_0.py'], check=True)
    except Exception as e:
        logger.error(f"进口处理程序运行失败: {e}")

def run_export_processor():
    """运行出口处理程序"""
    try:
        # 使用主控制器启动出口处理程序
        subprocess.run([sys.executable, 'OutputAutoRWwithSend_3_0.py'], check=True)
    except Exception as e:
        logger.error(f"出口处理程序运行失败: {e}")

def create_html_templates():
    """创建HTML模板文件"""
    # 确保模板目录存在
    os.makedirs('templates', exist_ok=True)
    
    # 主页面模板（使用之前提供的完整HTML代码）
    # 这里为了节省空间，不重复写入完整的HTML
    # 使用之前提供的完整HTML代码
    
    index_html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>舱单自动处理系统 - 管理界面</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.1/font/bootstrap-icons.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            background-color: #f8f9fa;
            font-family: 'Microsoft YaHei', Arial, sans-serif;
        }
        .navbar-brand {
            font-weight: bold;
        }
        .card {
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .card-header {
            background-color: #fff;
            border-bottom: 2px solid #e9ecef;
            font-weight: bold;
        }
        .status-badge {
            padding: 5px 10px;
            border-radius: 20px;
            font-size: 0.9em;
        }
        .status-running {
            background-color: #d4edda;
            color: #155724;
        }
        .status-stopped {
            background-color: #f8d7da;
            color: #721c24;
        }
        .stat-card {
            text-align: center;
            padding: 20px;
        }
        .stat-number {
            font-size: 2.5em;
            font-weight: bold;
            margin: 10px 0;
        }
        .stat-label {
            color: #6c757d;
            font-size: 0.9em;
        }
        .import-color {
            color: #3498db;
        }
        .export-color {
            color: #2ecc71;
        }
        .chart-container {
            position: relative;
            height: 300px;
            width: 100%;
        }
        .nav-tabs .nav-link {
            color: #495057;
            font-weight: 500;
        }
        .nav-tabs .nav-link.active {
            color: #0d6efd;
            font-weight: bold;
        }
        .table th {
            background-color: #f8f9fa;
            font-weight: 600;
        }
        .keyword-badge {
            background-color: #e9ecef;
            color: #495057;
            margin-right: 5px;
            margin-bottom: 5px;
        }
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(255,255,255,0.8);
            z-index: 9999;
            display: flex;
            justify-content: center;
            align-items: center;
        }
    </style>
</head>
<body>
    <!-- 导航栏 -->
    <nav class="navbar navbar-expand-lg navbar-light bg-white shadow-sm">
        <div class="container-fluid">
            <a class="navbar-brand" href="#">
                <i class="bi bi-envelope-check me-2"></i>
                舱单自动处理系统
            </a>
            <div class="navbar-nav">
                <a class="nav-link active" href="#" onclick="showTab('dashboard')">
                    <i class="bi bi-speedometer2 me-1"></i>仪表盘
                </a>
                <a class="nav-link" href="#" onclick="showTab('database')">
                    <i class="bi bi-database me-1"></i>数据库
                </a>
                <a class="nav-link" href="#" onclick="showTab('statistics')">
                    <i class="bi bi-bar-chart me-1"></i>统计
                </a>
                <a class="nav-link" href="#" onclick="showTab('logs')">
                    <i class="bi bi-journal-text me-1"></i>日志
                </a>
                <a class="nav-link" href="#" onclick="showTab('tools')">
                    <i class="bi bi-tools me-1"></i>工具
                </a>
                <!-- 添加系统配置链接 -->
                <a class="nav-link" href="/config">
                    <i class="bi bi-gear me-1"></i>系统配置
                </a>
            </div>
        </div>
    </nav>

    <!-- 主要内容区 -->
    <div class="container-fluid mt-4">
        <!-- 仪表盘 -->
        <div id="dashboard-tab" class="tab-content">
            <!-- 系统状态卡片 -->
            <div class="row mb-4">
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <i class="bi bi-cpu me-2"></i>系统状态
                        </div>
                        <div class="card-body">
                            <div class="row">
                                <div class="col-md-6">
                                    <div class="d-flex align-items-center mb-3">
                                        <div class="me-3">
                                            <i class="bi bi-download import-color" style="font-size: 1.5em;"></i>
                                        </div>
                                        <div>
                                            <div>进口舱单处理</div>
                                            <div id="import-status" class="status-badge status-stopped">已停止</div>
                                        </div>
                                    </div>
                                </div>
                                <div class="col-md-6">
                                    <div class="d-flex align-items-center mb-3">
                                        <div class="me-3">
                                            <i class="bi bi-upload export-color" style="font-size: 1.5em;"></i>
                                        </div>
                                        <div>
                                            <div>出口舱单处理</div>
                                            <div id="export-status" class="status-badge status-stopped">已停止</div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="text-muted small" id="last-check">
                                最后检查: --
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <i class="bi bi-calendar-check me-2"></i>今日统计
                        </div>
                        <div class="card-body">
                            <div class="row text-center">
                                <div class="col-6">
                                    <div class="stat-card">
                                        <div class="import-color stat-number" id="today-import">0</div>
                                        <div class="stat-label">今日进口</div>
                                    </div>
                                </div>
                                <div class="col-6">
                                    <div class="stat-card">
                                        <div class="export-color stat-number" id="today-export">0</div>
                                        <div class="stat-label">今日出口</div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 数据库统计 -->
            <div class="row mb-4">
                <div class="col-md-12">
                    <div class="card">
                        <div class="card-header">
                            <i class="bi bi-database me-2"></i>数据库统计
                        </div>
                        <div class="card-body">
                            <div class="row text-center">
                                <div class="col-md-3">
                                    <div class="stat-card">
                                        <div class="import-color stat-number" id="total-import">0</div>
                                        <div class="stat-label">进口总记录</div>
                                    </div>
                                </div>
                                <div class="col-md-3">
                                    <div class="stat-card">
                                        <div class="export-color stat-number" id="total-export">0</div>
                                        <div class="stat-label">出口总记录</div>
                                    </div>
                                </div>
                                <div class="col-md-3">
                                    <div class="stat-card">
                                        <div class="text-primary stat-number" id="total-all">0</div>
                                        <div class="stat-label">总记录数</div>
                                    </div>
                                </div>
                                <div class="col-md-3">
                                    <button class="btn btn-primary mt-4" onclick="showTab('database')">
                                        <i class="bi bi-search me-1"></i>查看详情
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 统计图表 -->
            <div class="row">
                <div class="col-md-12">
                    <div class="card">
                        <div class="card-header d-flex justify-content-between align-items-center">
                            <div>
                                <i class="bi bi-bar-chart me-2"></i>处理趋势
                            </div>
                            <div>
                                <select class="form-select form-select-sm" style="width: auto;" onchange="updateChartDays(this.value)">
                                    <option value="7">最近7天</option>
                                    <option value="30" selected>最近30天</option>
                                    <option value="90">最近90天</option>
                                </select>
                            </div>
                        </div>
                        <div class="card-body">
                            <div class="chart-container">
                                <canvas id="trendChart"></canvas>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 数据库页面 -->
        <div id="database-tab" class="tab-content" style="display: none;">
            <div class="card">
                <div class="card-header">
                    <ul class="nav nav-tabs card-header-tabs">
                        <li class="nav-item">
                            <a class="nav-link active" href="#" onclick="showDatabaseTab('import')">进口数据库</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="#" onclick="showDatabaseTab('export')">出口数据库</a>
                        </li>
                    </ul>
                </div>
                <div class="card-body">
                    <!-- 搜索和筛选 -->
                    <div class="row mb-3">
                        <div class="col-md-3">
                            <label class="form-label">开始日期</label>
                            <input type="date" class="form-control" id="db-start-date">
                        </div>
                        <div class="col-md-3">
                            <label class="form-label">结束日期</label>
                            <input type="date" class="form-control" id="db-end-date">
                        </div>
                        <div class="col-md-4">
                            <label class="form-label">关键词</label>
                            <input type="text" class="form-control" id="db-keywords" placeholder="多个关键词用逗号分隔">
                        </div>
                        <div class="col-md-2 d-flex align-items-end">
                            <button class="btn btn-primary w-100" onclick="loadDatabase()">
                                <i class="bi bi-search me-1"></i>搜索
                            </button>
                        </div>
                    </div>

                    <!-- 进口数据库表格 -->
                    <div id="import-database" class="database-table">
                        <div class="d-flex justify-content-between align-items-center mb-3">
                            <h5>进口舱单记录</h5>
                            <button class="btn btn-sm btn-outline-primary" onclick="exportCSV('import')">
                                <i class="bi bi-download me-1"></i>导出CSV
                            </button>
                        </div>
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th>处理时间</th>
                                        <th>发件人</th>
                                        <th>主题</th>
                                        <th>关键词</th>
                                        <th>箱数</th>
                                        <th>附件</th>
                                    </tr>
                                </thead>
                                <tbody id="import-table-body">
                                    <!-- 数据将通过JavaScript填充 -->
                                </tbody>
                            </table>
                        </div>
                        <div class="d-flex justify-content-between align-items-center mt-3">
                            <div id="import-table-info">显示 0 条记录</div>
                            <nav>
                                <ul class="pagination pagination-sm" id="import-pagination">
                                    <!-- 分页将通过JavaScript生成 -->
                                </ul>
                            </nav>
                        </div>
                    </div>

                    <!-- 出口数据库表格 -->
                    <div id="export-database" class="database-table" style="display: none;">
                        <div class="d-flex justify-content-between align-items-center mb-3">
                            <h5>出口舱单记录</h5>
                            <button class="btn btn-sm btn-outline-primary" onclick="exportCSV('export')">
                                <i class="bi bi-download me-1"></i>导出CSV
                            </button>
                        </div>
                        <div class="table-responsive">
                            <table class="table table-hover">
                                <thead>
                                    <tr>
                                        <th>处理时间</th>
                                        <th>发件人</th>
                                        <th>主题</th>
                                        <th>关键词</th>
                                        <th>箱数</th>
                                        <th>附件</th>
                                    </tr>
                                </thead>
                                <tbody id="export-table-body">
                                    <!-- 数据将通过JavaScript填充 -->
                                </tbody>
                            </table>
                        </div>
                        <div class="d-flex justify-content-between align-items-center mt-3">
                            <div id="export-table-info">显示 0 条记录</div>
                            <nav>
                                <ul class="pagination pagination-sm" id="export-pagination">
                                    <!-- 分页将通过JavaScript生成 -->
                                </ul>
                            </nav>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 统计页面 -->
        <div id="statistics-tab" class="tab-content" style="display: none;">
            <div class="card">
                <div class="card-header">
                    <ul class="nav nav-tabs card-header-tabs">
                        <li class="nav-item">
                            <a class="nav-link active" href="#" onclick="showStatisticsTab('keywords')">关键词统计</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="#" onclick="showStatisticsTab('charts')">图表分析</a>
                        </li>
                    </ul>
                </div>
                <div class="card-body">
                    <!-- 关键词统计 -->
                    <div id="keywords-statistics">
                        <div class="row mb-3">
                            <div class="col-md-3">
                                <label class="form-label">开始日期</label>
                                <input type="date" class="form-control" id="stats-start-date">
                            </div>
                            <div class="col-md-3">
                                <label class="form-label">结束日期</label>
                                <input type="date" class="form-control" id="stats-end-date">
                            </div>
                            <div class="col-md-3 d-flex align-items-end">
                                <button class="btn btn-primary" onclick="loadKeywordStatistics()">
                                    <i class="bi bi-calculator me-1"></i>统计
                                </button>
                            </div>
                        </div>

                        <div class="row">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <i class="bi bi-download import-color me-2"></i>进口关键词统计
                                    </div>
                                    <div class="card-body">
                                        <div id="import-keywords-list">
                                            <div class="text-center text-muted py-3">
                                                请选择日期范围进行统计
                                            </div>
                                        </div>
                                        <div class="mt-3">
                                            <strong>总计: </strong>
                                            <span id="import-total-count">0</span> 个附件
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">
                                        <i class="bi bi-upload export-color me-2"></i>出口关键词统计
                                    </div>
                                    <div class="card-body">
                                        <div id="export-keywords-list">
                                            <div class="text-center text-muted py-3">
                                                请选择日期范围进行统计
                                            </div>
                                        </div>
                                        <div class="mt-3">
                                            <strong>总计: </strong>
                                            <span id="export-total-count">0</span> 个附件
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class="row mt-4">
                            <div class="col-md-12">
                                <div class="card">
                                    <div class="card-header">
                                        <i class="bi bi-list-check me-2"></i>关键词汇总
                                    </div>
                                    <div class="card-body">
                                        <div id="all-keywords-list">
                                            <div class="text-center text-muted py-3">
                                                所有出现过的关键词将显示在这里
                                            </div>
                                        </div>
                                        <div class="mt-3">
                                            <button class="btn btn-outline-primary" onclick="exportKeywordChart()">
                                                <i class="bi bi-image me-1"></i>导出图表
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- 图表分析 -->
                    <div id="charts-statistics" style="display: none;">
                        <div class="row">
                            <div class="col-md-12">
                                <div class="card">
                                    <div class="card-header">
                                        <i class="bi bi-bar-chart me-2"></i>处理趋势图表
                                    </div>
                                    <div class="card-body">
                                        <div class="row mb-3">
                                            <div class="col-md-3">
                                                <select class="form-select" onchange="loadDailyChart()" id="chart-days">
                                                    <option value="7">最近7天</option>
                                                    <option value="30" selected>最近30天</option>
                                                    <option value="90">最近90天</option>
                                                </select>
                                            </div>
                                            <div class="col-md-3">
                                                <select class="form-select" onchange="loadDailyChart()" id="chart-type">
                                                    <option value="all">进出口合计</option>
                                                    <option value="import">仅进口</option>
                                                    <option value="export">仅出口</option>
                                                </select>
                                            </div>
                                            <div class="col-md-3">
                                                <button class="btn btn-outline-primary" onclick="exportDailyChart()">
                                                    <i class="bi bi-download me-1"></i>导出图表
                                                </button>
                                            </div>
                                        </div>
                                        <div class="chart-container">
                                            <canvas id="dailyChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class="row mt-4">
                            <div class="col-md-12">
                                <div class="card">
                                    <div class="card-header">
                                        <i class="bi bi-pie-chart me-2"></i>关键词分布图表
                                    </div>
                                    <div class="card-body">
                                        <div class="row mb-3">
                                            <div class="col-md-3">
                                                <input type="date" class="form-control" id="keyword-chart-start">
                                            </div>
                                            <div class="col-md-3">
                                                <input type="date" class="form-control" id="keyword-chart-end">
                                            </div>
                                            <div class="col-md-3">
                                                <button class="btn btn-primary" onclick="loadKeywordChart()">
                                                    <i class="bi bi-refresh me-1"></i>更新图表
                                                </button>
                                            </div>
                                        </div>
                                        <div class="chart-container">
                                            <img id="keyword-chart-img" src="" alt="关键词图表" style="width: 100%;">
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 日志页面 -->
        <div id="logs-tab" class="tab-content" style="display: none;">
            <div class="card">
                <div class="card-header">
                    <ul class="nav nav-tabs card-header-tabs">
                        <li class="nav-item">
                            <a class="nav-link active" href="#" onclick="showLogsTab('import')">进口日志</a>
                        </li>
                        <li class="nav-item">
                            <a class="nav-link" href="#" onclick="showLogsTab('export')">出口日志</a>
                        </li>
                    </ul>
                </div>
                <div class="card-body">
                    <div class="row mb-3">
                        <div class="col-md-2">
                            <select class="form-select" id="log-lines" onchange="loadLogs()">
                                <option value="50">最近50条</option>
                                <option value="100" selected>最近100条</option>
                                <option value="200">最近200条</option>
                                <option value="500">最近500条</option>
                            </select>
                        </div>
                    </div>

                    <!-- 进口日志 -->
                    <div id="import-logs" class="logs-table">
                        <div class="table-responsive">
                            <table class="table table-sm table-hover">
                                <thead>
                                    <tr>
                                        <th>时间</th>
                                        <th>邮件ID</th>
                                        <th>发件人</th>
                                        <th>主题</th>
                                        <th>关键词</th>
                                        <th>状态</th>
                                    </tr>
                                </thead>
                                <tbody id="import-logs-body">
                                    <!-- 日志数据将通过JavaScript填充 -->
                                </tbody>
                            </table>
                        </div>
                    </div>

                    <!-- 出口日志 -->
                    <div id="export-logs" class="logs-table" style="display: none;">
                        <div class="table-responsive">
                            <table class="table table-sm table-hover">
                                <thead>
                                    <tr>
                                        <th>时间</th>
                                        <th>邮件ID</th>
                                        <th>发件人</th>
                                        <th>主题</th>
                                        <th>关键词</th>
                                        <th>状态</th>
                                    </tr>
                                </thead>
                                <tbody id="export-logs-body">
                                    <!-- 日志数据将通过JavaScript填充 -->
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- 工具页面 -->
        <div id="tools-tab" class="tab-content" style="display: none;">
            <div class="row">
                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <i class="bi bi-chat-dots me-2"></i>短信测试
                        </div>
                        <div class="card-body">
                            <div class="mb-3">
                                <label class="form-label">选择测试类型</label>
                                <select class="form-select" id="sms-type">
                                    <option value="import">进口舱单短信</option>
                                    <option value="export">出口舱单短信</option>
                                </select>
                            </div>
                            <div class="mb-3">
                                <label class="form-label">测试内容</label>
                                <textarea class="form-control" id="sms-content" rows="3" readonly>
【天津港集装箱码头有限公司】舱单处理程序测试短信
                                </textarea>
                            </div>
                            <button class="btn btn-primary" onclick="testSMS()">
                                <i class="bi bi-send me-1"></i>发送测试短信
                            </button>
                            <div id="sms-result" class="mt-3"></div>
                        </div>
                    </div>
                </div>

                <div class="col-md-6">
                    <div class="card">
                        <div class="card-header">
                            <i class="bi bi-gear me-2"></i>系统控制
                        </div>
                        <div class="card-body">
                            <div class="mb-3">
                                <label class="form-label">系统操作</label>
                                <div class="d-grid gap-2">
                                    <button class="btn btn-success" onclick="startSystem()">
                                        <i class="bi bi-play-circle me-1"></i>启动处理系统
                                    </button>
                                    <button class="btn btn-warning" onclick="syncHistory()">
                                        <i class="bi bi-arrow-repeat me-1"></i>同步历史邮件
                                    </button>
                                    <button class="btn btn-info" onclick="checkDatabase()">
                                        <i class="bi bi-database-check me-1"></i>检查数据库
                                    </button>
                                </div>
                            </div>
                            <div id="system-control-result" class="mt-3"></div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="row mt-4">
                <div class="col-md-12">
                    <div class="card">
                        <div class="card-header">
                            <i class="bi bi-key me-2"></i>关键词管理
                        </div>
                        <div class="card-body">
                            <div class="row">
                                <div class="col-md-6">
                                    <h6>进口关键词</h6>
                                    <div id="import-keywords-display" class="mb-3">
                                        <!-- 关键词将通过JavaScript显示 -->
                                    </div>
                                </div>
                                <div class="col-md-6">
                                    <h6>出口关键词</h6>
                                    <div id="export-keywords-display" class="mb-3">
                                        <!-- 关键词将通过JavaScript显示 -->
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- 加载遮罩 -->
    <div id="loading-overlay" class="loading-overlay" style="display: none;">
        <div class="text-center">
            <div class="spinner-border text-primary" role="status">
                <span class="visually-hidden">加载中...</span>
            </div>
            <div class="mt-2">加载中...</div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // 全局变量
        let currentTab = 'dashboard';
        let currentDatabaseTab = 'import';
        let currentStatisticsTab = 'keywords';
        let currentLogsTab = 'import';
        let trendChart = null;
        let dailyChart = null;

        // 页面加载完成后执行
        document.addEventListener('DOMContentLoaded', function() {
            // 设置默认日期
            const today = new Date().toISOString().split('T')[0];
            const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
            
            document.getElementById('db-start-date').value = weekAgo;
            document.getElementById('db-end-date').value = today;
            document.getElementById('stats-start-date').value = weekAgo;
            document.getElementById('stats-end-date').value = today;
            document.getElementById('keyword-chart-start').value = weekAgo;
            document.getElementById('keyword-chart-end').value = today;

            // 初始化系统
            loadSystemStatus();
            loadDatabase();
            loadKeywordStatistics();
            loadDailyChart();
            loadLogs();
            loadKeywords();

            // 定时刷新系统状态
            setInterval(loadSystemStatus, 10000);
            setInterval(loadKeywordChart, 30000); // 30秒更新一次关键词图表
        });

        // 显示标签页
        function showTab(tabName) {
            // 隐藏所有标签页
            document.querySelectorAll('.tab-content').forEach(tab => {
                tab.style.display = 'none';
            });
            
            // 移除所有导航链接的激活状态
            document.querySelectorAll('.navbar-nav .nav-link').forEach(link => {
                link.classList.remove('active');
            });
            
            // 显示选中的标签页
            document.getElementById(tabName + '-tab').style.display = 'block';
            
            // 激活对应的导航链接
            document.querySelector(`.navbar-nav .nav-link[onclick*="${tabName}"]`).classList.add('active');
            
            currentTab = tabName;
            
            // 如果是统计页，确保图表正确显示
            if (tabName === 'statistics') {
                setTimeout(() => {
                    if (currentStatisticsTab === 'charts') {
                        loadDailyChart();
                        loadKeywordChart();
                    }
                }, 100);
            }
        }

        // 显示数据库子标签页
        function showDatabaseTab(tabName) {
            currentDatabaseTab = tabName;
            
            // 更新标签页激活状态
            document.querySelectorAll('#database-tab .nav-link').forEach(link => {
                link.classList.remove('active');
            });
            event.target.classList.add('active');
            
            // 显示对应的数据库表格
            document.querySelectorAll('.database-table').forEach(table => {
                table.style.display = 'none';
            });
            document.getElementById(tabName + '-database').style.display = 'block';
            
            // 加载对应的数据库数据
            loadDatabase();
        }

        // 显示统计子标签页
        function showStatisticsTab(tabName) {
            currentStatisticsTab = tabName;
            
            // 更新标签页激活状态
            document.querySelectorAll('#statistics-tab .nav-link').forEach(link => {
                link.classList.remove('active');
            });
            event.target.classList.add('active');
            
            // 显示对应的统计内容
            document.getElementById('keywords-statistics').style.display = tabName === 'keywords' ? 'block' : 'none';
            document.getElementById('charts-statistics').style.display = tabName === 'charts' ? 'block' : 'none';
            
            // 如果是图表页，加载图表
            if (tabName === 'charts') {
                setTimeout(() => {
                    loadDailyChart();
                    loadKeywordChart();
                }, 100);
            }
        }

        // 显示日志子标签页
        function showLogsTab(tabName) {
            currentLogsTab = tabName;
            
            // 更新标签页激活状态
            document.querySelectorAll('#logs-tab .nav-link').forEach(link => {
                link.classList.remove('active');
            });
            event.target.classList.add('active');
            
            // 显示对应的日志表格
            document.querySelectorAll('.logs-table').forEach(table => {
                table.style.display = 'none';
            });
            document.getElementById(tabName + '-logs').style.display = 'block';
            
            // 加载对应的日志数据
            loadLogs();
        }

        // 加载系统状态
        function loadSystemStatus() {
            fetch('/api/system/status')
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        const status = data.data.system;
                        const db = data.data.database;
                        
                        // 更新系统状态
                        const importStatus = document.getElementById('import-status');
                        const exportStatus = document.getElementById('export-status');
                        
                        if (status.import_running) {
                            importStatus.textContent = '运行中';
                            importStatus.className = 'status-badge status-running';
                        } else {
                            importStatus.textContent = '已停止';
                            importStatus.className = 'status-badge status-stopped';
                        }
                        
                        if (status.export_running) {
                            exportStatus.textContent = '运行中';
                            exportStatus.className = 'status-badge status-running';
                        } else {
                            exportStatus.textContent = '已停止';
                            exportStatus.className = 'status-badge status-stopped';
                        }
                        
                        // 更新最后检查时间
                        document.getElementById('last-check').textContent = `最后检查: ${status.last_check || '--'}`;
                        
                        // 更新数据库统计
                        document.getElementById('today-import').textContent = db.today_import;
                        document.getElementById('today-export').textContent = db.today_export;
                        document.getElementById('total-import').textContent = db.import_total;
                        document.getElementById('total-export').textContent = db.export_total;
                        document.getElementById('total-all').textContent = db.import_total + db.export_total;
                    }
                })
                .catch(error => {
                    console.error('加载系统状态失败:', error);
                });
        }

        // 加载数据库数据
        function loadDatabase(page = 1) {
            showLoading();
            
            const startDate = document.getElementById('db-start-date').value;
            const endDate = document.getElementById('db-end-date').value;
            const keywords = document.getElementById('db-keywords').value;
            
            const endpoint = currentDatabaseTab === 'import' ? '/api/import/database' : '/api/export/database';
            
            fetch(`${endpoint}?page=${page}&start_date=${startDate}&end_date=${endDate}&keywords=${keywords}`)
                .then(response => response.json())
                .then(data => {
                    hideLoading();
                    
                    if (data.success) {
                        renderDatabaseTable(data, currentDatabaseTab);
                    } else {
                        alert('加载数据失败: ' + (data.error || '未知错误'));
                    }
                })
                .catch(error => {
                    hideLoading();
                    console.error('加载数据库失败:', error);
                    alert('加载数据库失败: ' + error.message);
                });
        }

        // 渲染数据库表格
        function renderDatabaseTable(data, type) {
            const tableBody = document.getElementById(`${type}-table-body`);
            const tableInfo = document.getElementById(`${type}-table-info`);
            const pagination = document.getElementById(`${type}-pagination`);
            
            // 清空表格
            tableBody.innerHTML = '';
            
            if (data.data.length === 0) {
                tableBody.innerHTML = `
                    <tr>
                        <td colspan="6" class="text-center text-muted py-4">
                            没有找到匹配的记录
                        </td>
                    </tr>
                `;
                tableInfo.textContent = '显示 0 条记录';
                pagination.innerHTML = '';
                return;
            }
            
            // 填充数据
            data.data.forEach(item => {
                const row = document.createElement('tr');
                
                // 格式化时间
                const processedDate = item.processed_date ? item.processed_date.split(' ')[0] : '--';
                
                // 截长发件人和主题
                const sender = item.sender && item.sender.length > 30 ? 
                    item.sender.substring(0, 30) + '...' : item.sender;
                const subject = item.subject && item.subject.length > 40 ? 
                    item.subject.substring(0, 40) + '...' : item.subject;
                
                // 关键词标签
                let keywordBadges = '';
                if (item.matched_keywords) {
                    const keywords = item.matched_keywords.split(',');
                    keywords.forEach(keyword => {
                        if (keyword.trim()) {
                            keywordBadges += `<span class="badge keyword-badge">${keyword.trim()}</span>`;
                        }
                    });
                }
                
                row.innerHTML = `
                    <td>${processedDate}</td>
                    <td title="${item.sender || ''}">${sender || '--'}</td>
                    <td title="${item.subject || ''}">${subject || '--'}</td>
                    <td>${keywordBadges || '--'}</td>
                    <td>${item.container_count || 0}</td>
                    <td title="${item.attachment_names || ''}">
                        ${item.txt_attachment ? item.txt_attachment.substring(0, 20) + '...' : '--'}
                    </td>
                `;
                
                tableBody.appendChild(row);
            });
            
            // 更新表格信息
            const start = (data.page - 1) * data.page_size + 1;
            const end = Math.min(data.page * data.page_size, data.total);
            tableInfo.textContent = `显示 ${start}-${end} 条记录，共 ${data.total} 条`;
            
            // 生成分页
            renderPagination(pagination, data.page, data.total_pages, (newPage) => {
                loadDatabase(newPage);
            });
        }

        // 生成分页控件
        function renderPagination(container, currentPage, totalPages, onClick) {
            container.innerHTML = '';
            
            if (totalPages <= 1) return;
            
            // 上一页按钮
            const prevLi = document.createElement('li');
            prevLi.className = `page-item ${currentPage === 1 ? 'disabled' : ''}`;
            prevLi.innerHTML = `<a class="page-link" href="#" ${currentPage === 1 ? 'tabindex="-1"' : ''}>上一页</a>`;
            if (currentPage > 1) {
                prevLi.querySelector('a').addEventListener('click', (e) => {
                    e.preventDefault();
                    onClick(currentPage - 1);
                });
            }
            container.appendChild(prevLi);
            
            // 页码按钮
            const startPage = Math.max(1, currentPage - 2);
            const endPage = Math.min(totalPages, startPage + 4);
            
            for (let i = startPage; i <= endPage; i++) {
                const pageLi = document.createElement('li');
                pageLi.className = `page-item ${i === currentPage ? 'active' : ''}`;
                pageLi.innerHTML = `<a class="page-link" href="#">${i}</a>`;
                
                pageLi.querySelector('a').addEventListener('click', (e) => {
                    e.preventDefault();
                    if (i !== currentPage) {
                        onClick(i);
                    }
                });
                
                container.appendChild(pageLi);
            }
            
            // 下一页按钮
            const nextLi = document.createElement('li');
            nextLi.className = `page-item ${currentPage === totalPages ? 'disabled' : ''}`;
            nextLi.innerHTML = `<a class="page-link" href="#" ${currentPage === totalPages ? 'tabindex="-1"' : ''}>下一页</a>`;
            if (currentPage < totalPages) {
                nextLi.querySelector('a').addEventListener('click', (e) => {
                    e.preventDefault();
                    onClick(currentPage + 1);
                });
            }
            container.appendChild(nextLi);
        }

        // 加载关键词统计
        function loadKeywordStatistics() {
            showLoading();
            
            const startDate = document.getElementById('stats-start-date').value;
            const endDate = document.getElementById('stats-end-date').value;
            
            fetch(`/api/statistics/keywords?start_date=${startDate}&end_date=${endDate}`)
                .then(response => response.json())
                .then(data => {
                    hideLoading();
                    
                    if (data.success) {
                        renderKeywordStatistics(data.data);
                    } else {
                        alert('加载统计失败: ' + (data.error || '未知错误'));
                    }
                })
                .catch(error => {
                    hideLoading();
                    console.error('加载关键词统计失败:', error);
                    alert('加载统计失败: ' + error.message);
                });
        }

        // 渲染关键词统计
        function renderKeywordStatistics(data) {
            // 进口关键词
            const importList = document.getElementById('import-keywords-list');
            importList.innerHTML = '';
            
            if (Object.keys(data.import).length === 0) {
                importList.innerHTML = '<div class="text-center text-muted py-3">没有找到进口关键词统计</div>';
            } else {
                let importHtml = '';
                for (const [keyword, count] of Object.entries(data.import)) {
                    importHtml += `
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <span>${keyword}</span>
                            <span class="badge bg-primary">${count}</span>
                        </div>
                    `;
                }
                importList.innerHTML = importHtml;
            }
            
            document.getElementById('import-total-count').textContent = data.total_import;
            
            // 出口关键词
            const exportList = document.getElementById('export-keywords-list');
            exportList.innerHTML = '';
            
            if (Object.keys(data.export).length === 0) {
                exportList.innerHTML = '<div class="text-center text-muted py-3">没有找到出口关键词统计</div>';
            } else {
                let exportHtml = '';
                for (const [keyword, count] of Object.entries(data.export)) {
                    exportHtml += `
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <span>${keyword}</span>
                            <span class="badge bg-success">${count}</span>
                        </div>
                    `;
                }
                exportList.innerHTML = exportHtml;
            }
            
            document.getElementById('export-total-count').textContent = data.total_export;
            
            // 所有关键词
            const allKeywordsList = document.getElementById('all-keywords-list');
            allKeywordsList.innerHTML = '';
            
            if (data.all_keywords.length === 0) {
                allKeywordsList.innerHTML = '<div class="text-center text-muted py-3">没有找到关键词</div>';
            } else {
                let allKeywordsHtml = '';
                data.all_keywords.forEach(keyword => {
                    const importCount = data.import[keyword] || 0;
                    const exportCount = data.export[keyword] || 0;
                    const totalCount = importCount + exportCount;
                    
                    allKeywordsHtml += `
                        <div class="d-flex justify-content-between align-items-center mb-2">
                            <div>
                                <span class="me-3">${keyword}</span>
                                <span class="badge import-color me-1">进口: ${importCount}</span>
                                <span class="badge export-color">出口: ${exportCount}</span>
                            </div>
                            <span class="badge bg-secondary">总计: ${totalCount}</span>
                        </div>
                    `;
                });
                allKeywordsList.innerHTML = allKeywordsHtml;
            }
        }

        // 加载每日统计图表
        function loadDailyChart() {
            const days = document.getElementById('chart-days').value;
            const type = document.getElementById('chart-type').value;
            
            fetch(`/api/chart/daily?days=${days}&type=${type}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success && data.image) {
                        // 创建图片元素
                        const img = new Image();
                        img.src = data.image;
                        img.onload = function() {
                            const canvas = document.getElementById('dailyChart');
                            const ctx = canvas.getContext('2d');
                            
                            // 设置canvas尺寸
                            canvas.width = img.width;
                            canvas.height = img.height;
                            
                            // 绘制图片
                            ctx.drawImage(img, 0, 0);
                        };
                    }
                })
                .catch(error => {
                    console.error('加载图表失败:', error);
                });
        }

        // 加载关键词图表
        function loadKeywordChart() {
            const startDate = document.getElementById('keyword-chart-start').value;
            const endDate = document.getElementById('keyword-chart-end').value;
            
            fetch(`/api/chart/keywords?start_date=${startDate}&end_date=${endDate}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success && data.image) {
                        document.getElementById('keyword-chart-img').src = data.image;
                    }
                })
                .catch(error => {
                    console.error('加载关键词图表失败:', error);
                });
        }

        // 加载日志
        function loadLogs() {
            const lines = document.getElementById('log-lines').value;
            const endpoint = currentLogsTab === 'import' ? '/api/logs/import' : '/api/logs/export';
            
            fetch(`${endpoint}?lines=${lines}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        renderLogsTable(data.data, currentLogsTab);
                    }
                })
                .catch(error => {
                    console.error('加载日志失败:', error);
                });
        }

        // 渲染日志表格
        function renderLogsTable(logs, type) {
            const tableBody = document.getElementById(`${type}-logs-body`);
            tableBody.innerHTML = '';
            
            if (logs.length === 0) {
                tableBody.innerHTML = `
                    <tr>
                        <td colspan="6" class="text-center text-muted py-4">
                            没有找到日志记录
                        </td>
                    </tr>
                `;
                return;
            }
            
            logs.forEach(log => {
                const row = document.createElement('tr');
                
                // 格式化时间
                const time = log.timestamp ? log.timestamp.split(' ')[1] : '--';
                
                // 状态标签
                let statusBadge = '';
                if (log.excel_sent) {
                    statusBadge = '<span class="badge bg-success">已发送</span>';
                } else if (log.has_keyword) {
                    statusBadge = '<span class="badge bg-warning">有关键词</span>';
                } else {
                    statusBadge = '<span class="badge bg-secondary">无关键词</span>';
                }
                
                row.innerHTML = `
                    <td title="${log.timestamp || ''}">${time || '--'}</td>
                    <td title="${log.email_uid || ''}">${log.email_uid ? log.email_uid.substring(0, 10) + '...' : '--'}</td>
                    <td title="${log.sender || ''}">${log.sender && log.sender.length > 20 ? log.sender.substring(0, 20) + '...' : log.sender || '--'}</td>
                    <td title="${log.subject || ''}">${log.subject && log.subject.length > 30 ? log.subject.substring(0, 30) + '...' : log.subject || '--'}</td>
                    <td>${log.matched_keywords || '--'}</td>
                    <td>${statusBadge}</td>
                `;
                
                tableBody.appendChild(row);
            });
        }

        // 加载关键词设置
        function loadKeywords() {
            fetch('/api/settings/keywords')
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        renderKeywords(data.data);
                    }
                })
                .catch(error => {
                    console.error('加载关键词失败:', error);
                });
        }

        // 渲染关键词
        function renderKeywords(data) {
            // 进口关键词
            const importDisplay = document.getElementById('import-keywords-display');
            if (data.import && data.import.length > 0) {
                let importHtml = '';
                data.import.forEach(keyword => {
                    importHtml += `<span class="badge import-color me-1 mb-1">${keyword}</span>`;
                });
                importDisplay.innerHTML = importHtml;
            } else {
                importDisplay.innerHTML = '<div class="text-muted">未设置关键词</div>';
            }
            
            // 出口关键词
            const exportDisplay = document.getElementById('export-keywords-display');
            if (data.export && data.export.length > 0) {
                let exportHtml = '';
                data.export.forEach(keyword => {
                    exportHtml += `<span class="badge export-color me-1 mb-1">${keyword}</span>`;
                });
                exportDisplay.innerHTML = exportHtml;
            } else {
                exportDisplay.innerHTML = '<div class="text-muted">未设置关键词</div>';
            }
        }

        // 测试短信发送
        function testSMS() {
            const smsType = document.getElementById('sms-type').value;
            const resultDiv = document.getElementById('sms-result');
            
            resultDiv.innerHTML = '<div class="alert alert-info">正在发送测试短信...</div>';
            
            fetch(`/api/sms/test?type=${smsType}`)
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        resultDiv.innerHTML = `<div class="alert alert-success">${data.message}</div>`;
                    } else {
                        resultDiv.innerHTML = `<div class="alert alert-danger">发送失败: ${data.error}</div>`;
                    }
                })
                .catch(error => {
                    resultDiv.innerHTML = `<div class="alert alert-danger">请求失败: ${error.message}</div>`;
                    console.error('测试短信失败:', error);
                });
        }

        // 启动系统
        function startSystem() {
            const resultDiv = document.getElementById('system-control-result');
            
            resultDiv.innerHTML = '<div class="alert alert-info">正在启动系统...</div>';
            
            fetch('/api/system/start')
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        resultDiv.innerHTML = `<div class="alert alert-success">${data.message}</div>`;
                        // 刷新系统状态
                        setTimeout(loadSystemStatus, 2000);
                    } else {
                        resultDiv.innerHTML = `<div class="alert alert-danger">启动失败: ${data.error}</div>`;
                    }
                })
                .catch(error => {
                    resultDiv.innerHTML = `<div class="alert alert-danger">请求失败: ${error.message}</div>`;
                    console.error('启动系统失败:', error);
                });
        }

        // 同步历史邮件 - 修正版本
        function syncHistory() {
            const resultDiv = document.getElementById('system-control-result');
            resultDiv.innerHTML = '<div class="alert alert-info">正在同步历史邮件...</div>';
            
            fetch('/api/system/sync_history')
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        resultDiv.innerHTML = `<div class="alert alert-success">${data.message}</div>`;
                        if (data.output) {
                            resultDiv.innerHTML += `<pre class="mt-2" style="font-size: 0.8em;">${data.output}</pre>`;
                        }
                    } else {
                        resultDiv.innerHTML = `<div class="alert alert-danger">${data.message || data.error}</div>`;
                    }
                })
                .catch(error => {
                    resultDiv.innerHTML = `<div class="alert alert-danger">请求失败: ${error.message}</div>`;
                    console.error('同步历史邮件失败:', error);
                });
        }

        // 检查数据库 - 修正版本
        function checkDatabase() {
            const resultDiv = document.getElementById('system-control-result');
            resultDiv.innerHTML = '<div class="alert alert-info">正在检查数据库...</div>';
            
            fetch('/api/system/check_database')
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        let html = `<div class="alert alert-success">${data.message}</div>`;
                        if (data.issues && data.issues.length > 0) {
                            html += '<ul class="mt-2">';
                            data.issues.forEach(issue => {
                                html += `<li>${issue}</li>`;
                            });
                            html += '</ul>';
                        }
                        resultDiv.innerHTML = html;
                    } else {
                        resultDiv.innerHTML = `<div class="alert alert-danger">${data.message || data.error}</div>`;
                    }
                })
                .catch(error => {
                    resultDiv.innerHTML = `<div class="alert alert-danger">请求失败: ${error.message}</div>`;
                    console.error('检查数据库失败:', error);
                });
        }

        // 导出CSV
        function exportCSV(type) {
            const startDate = document.getElementById('db-start-date').value;
            const endDate = document.getElementById('db-end-date').value;
            
            window.open(`/api/export/csv?type=${type}&start_date=${startDate}&end_date=${endDate}`, '_blank');
        }

        // 导出图表
        function exportDailyChart() {
            const days = document.getElementById('chart-days').value;
            const type = document.getElementById('chart-type').value;
            
            const link = document.createElement('a');
            link.href = `/api/chart/daily?days=${days}&type=${type}&download=1`;
            link.download = `处理趋势_${days}天_${type}.png`;
            link.click();
        }

        function exportKeywordChart() {
            const startDate = document.getElementById('stats-start-date').value;
            const endDate = document.getElementById('stats-end-date').value;
            
            const link = document.createElement('a');
            link.href = `/api/chart/keywords?start_date=${startDate}&end_date=${endDate}&download=1`;
            link.download = `关键词统计_${startDate}_至_${endDate}.png`;
            link.click();
        }

        // 更新图表天数
        function updateChartDays(days) {
            loadDailyChart();
        }

        // 显示加载遮罩
        function showLoading() {
            document.getElementById('loading-overlay').style.display = 'flex';
        }

        // 隐藏加载遮罩
        function hideLoading() {
            document.getElementById('loading-overlay').style.display = 'none';
        }
    </script>
</body>
</html>'''
    
    # 写入主页面模板
    with open('templates/index.html', 'w', encoding='utf-8') as f:
        f.write(index_html)
    
    print("✅ HTML模板创建完成")

if __name__ == '__main__':
    # 创建必要的目录
    os.makedirs('static', exist_ok=True)
    os.makedirs('templates', exist_ok=True)
    
    # 确保数据库表存在
    ensure_database_exists()
    
    # 创建HTML模板
    create_html_templates()
    
    # 启动Web服务器
    print("舱单自动处理系统 - Web管理界面")
    print("=" * 60)
    print(f"访问地址: http://localhost:5000")
    print("=" * 60)
    
    app.run(host='0.0.0.0', port=5000, debug=False, use_reloader=False)