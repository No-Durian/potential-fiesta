"""
主控制器 - 修复编码问题的版本
"""
import sys
import io
import os

# 在程序开始时设置全局编码环境变量
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 强制设置标准输出和错误输出的编码为UTF-8
if sys.version_info >= (3, 7):
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
else:
    # Python 3.6及以下版本的兼容方案
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import threading
import time
import logging
import signal
import traceback

# 设置日志 - 使用UTF-8编码
def setup_logging():
    """设置日志配置"""
    # 清除现有的处理器
    root_logger = logging.getLogger()
    for handler in root_logger.handlers[:]:
        root_logger.removeHandler(handler)
    
    # 创建处理器
    handler = logging.StreamHandler(sys.stdout)
    # 创建一个能处理UTF-8的格式化器
    class UTF8Formatter(logging.Formatter):
        def format(self, record):
            try:
                result = super().format(record)
                return result
            except UnicodeEncodeError:
                # 如果遇到编码问题，使用安全的处理方式
                record.msg = record.msg.encode('utf-8', errors='replace').decode('utf-8', errors='replace')
                return super().format(record)
    
    formatter = UTF8Formatter('%(asctime)s - [%(threadName)s] - %(message)s')
    handler.setFormatter(formatter)
    
    # 设置编码
    if hasattr(handler.stream, 'reconfigure'):
        handler.stream.reconfigure(encoding='utf-8')
    
    root_logger.addHandler(handler)
    root_logger.setLevel(logging.INFO)
    
    # 禁用第三方库的日志
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    logging.getLogger('smtplib').setLevel(logging.WARNING)
    
    return root_logger

# 初始化日志
logger = setup_logging()

class ImportManifestProcessor:
    """进口舱单处理程序"""
    
    def __init__(self):
        self.thread = None
        self.running = False
        self.thread_name = "ImportProcessor"
        
    def start(self):
        """启动进口舱单处理程序"""
        if self.running:
            logger.info(f"{self.thread_name} 已经在运行")
            return
            
        self.running = True
        self.thread = threading.Thread(
            target=self._run_import_processor,
            name=self.thread_name,
            daemon=True
        )
        self.thread.start()
        logger.info(f"✅ {self.thread_name} 已启动")
        
    def stop(self):
        """停止进口舱单处理程序"""
        self.running = False
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=5)
        logger.info(f"🛑 {self.thread_name} 已停止")
        
    def send_manual_stop_notification(self):
        """发送手动停止通知"""
        try:
            # 动态导入，避免编码问题
            import importlib
            import sys
            import io
            
            # 临时重定向输出以捕获短信发送状态
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            
            try:
                # 创建一个StringIO来捕获输出
                output_capture = io.StringIO()
                sys.stdout = output_capture
                sys.stderr = output_capture
                
                # 导入模块
                sms_module = importlib.import_module('InputAutoRW_FullFunc_2_0')
                
                # 调用发送通知函数
                result = sms_module.send_exit_notification(is_manual=True)
                
                # 获取捕获的输出
                captured_output = output_capture.getvalue()
                
                # 输出到日志
                for line in captured_output.split('\n'):
                    if line.strip():
                        logger.info(f"📱 [进口] {line}")
                
                if result:
                    logger.info("✅ 进口舱单手动关闭短信通知发送成功")
                else:
                    logger.warning("⚠️ 进口舱单手动关闭短信通知发送失败")
                    
                return result
                
            finally:
                # 恢复标准输出
                sys.stdout = old_stdout
                sys.stderr = old_stderr
                
        except Exception as e:
            logger.error(f"❌ 发送进口舱单手动关闭通知失败: {e}")
            return False
        
    def _run_import_processor(self):
        """运行进口舱单处理程序的主逻辑"""
        try:
            # 设置环境变量
            os.environ['PYTHONIOENCODING'] = 'utf-8'
            
            # 动态导入进口舱单处理模块
            import importlib
            import sys
            import io
            
            # 临时重定向输出
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            
            try:
                # 创建一个能处理UTF-8的StringIO
                output_capture = io.StringIO()
                sys.stdout = output_capture
                sys.stderr = output_capture
                
                # 导入模块
                import_module = importlib.import_module('InputAutoRW_FullFunc_2_0')
                
                # 运行主函数
                import_module.main()
                
            except KeyboardInterrupt:
                logger.info(f"{self.thread_name} 被用户中断")
                try:
                    import_module.send_exit_notification(is_manual=True)
                except:
                    pass
            except Exception as e:
                logger.error(f"{self.thread_name} 异常退出: {e}")
                try:
                    import_module.send_exit_notification(str(e)[:100])
                except:
                    pass
            finally:
                # 恢复标准输出并处理捕获的输出
                sys.stdout = old_stdout
                sys.stderr = old_stderr
                
                # 处理捕获的输出
                captured_output = output_capture.getvalue()
                for line in captured_output.split('\n'):
                    if line.strip():
                        # 过滤掉一些调试信息
                        if 'DEBUG' not in line and 'urllib3' not in line:
                            logger.info(f"[进口] {line}")
                
        except Exception as e:
            logger.error(f"{self.thread_name} 启动失败: {e}")
        finally:
            self.running = False

class ExportManifestProcessor:
    """出口舱单处理程序"""
    
    def __init__(self):
        self.thread = None
        self.running = False
        self.thread_name = "ExportProcessor"
        
    def start(self):
        """启动出口舱单处理程序"""
        if self.running:
            logger.info(f"{self.thread_name} 已经在运行")
            return
            
        self.running = True
        self.thread = threading.Thread(
            target=self._run_export_processor,
            name=self.thread_name,
            daemon=True
        )
        self.thread.start()
        logger.info(f"✅ {self.thread_name} 已启动")
        
    def stop(self):
        """停止出口舱单处理程序"""
        self.running = False
        if self.thread and self.thread.is_alive():
            self.thread.join(timeout=5)
        logger.info(f"🛑 {self.thread_name} 已停止")
        
    def send_manual_stop_notification(self):
        """发送手动停止通知"""
        try:
            # 动态导入
            import importlib
            import sys
            import io
            
            # 临时重定向输出
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            
            try:
                output_capture = io.StringIO()
                sys.stdout = output_capture
                sys.stderr = output_capture
                
                # 导入模块
                sms_module = importlib.import_module('OutputAutoRWwithSend_3_0')
                
                # 调用发送通知函数
                result = sms_module.send_exit_notification(is_manual=True)
                
                # 获取捕获的输出
                captured_output = output_capture.getvalue()
                
                # 输出到日志
                for line in captured_output.split('\n'):
                    if line.strip():
                        logger.info(f"📱 [出口] {line}")
                
                if result:
                    logger.info("✅ 出口舱单手动关闭短信通知发送成功")
                else:
                    logger.warning("⚠️ 出口舱单手动关闭短信通知发送失败")
                    
                return result
                
            finally:
                # 恢复标准输出
                sys.stdout = old_stdout
                sys.stderr = old_stderr
                
        except Exception as e:
            logger.error(f"❌ 发送出口舱单手动关闭通知失败: {e}")
            return False
        
    def _run_export_processor(self):
        """运行出口舱单处理程序的主逻辑"""
        try:
            # 设置环境变量
            os.environ['PYTHONIOENCODING'] = 'utf-8'
            
            # 动态导入出口舱单处理模块
            import importlib
            import sys
            import io
            
            # 临时重定向输出
            old_stdout = sys.stdout
            old_stderr = sys.stderr
            
            try:
                # 创建一个能处理UTF-8的StringIO
                output_capture = io.StringIO()
                sys.stdout = output_capture
                sys.stderr = output_capture
                
                # 导入模块
                export_module = importlib.import_module('OutputAutoRWwithSend_3_0')
                
                # 运行主函数
                export_module.main()
                
            except KeyboardInterrupt:
                logger.info(f"{self.thread_name} 被用户中断")
                try:
                    export_module.send_exit_notification(is_manual=True)
                except:
                    pass
            except Exception as e:
                logger.error(f"{self.thread_name} 异常退出: {e}")
                try:
                    export_module.send_exit_notification(str(e)[:100])
                except:
                    pass
            finally:
                # 恢复标准输出并处理捕获的输出
                sys.stdout = old_stdout
                sys.stderr = old_stderr
                
                # 处理捕获的输出
                captured_output = output_capture.getvalue()
                for line in captured_output.split('\n'):
                    if line.strip():
                        # 过滤掉一些调试信息
                        if 'DEBUG' not in line and 'urllib3' not in line:
                            logger.info(f"[出口] {line}")
                
        except Exception as e:
            logger.error(f"{self.thread_name} 启动失败: {e}")
        finally:
            self.running = False

class MainController:
    """主控制器 - 管理所有处理程序"""
    
    def __init__(self):
        self.running = False
        self.import_processor = ImportManifestProcessor()
        self.export_processor = ExportManifestProcessor()
        
    def start_all(self):
        """启动所有处理程序"""
        if self.running:
            logger.info("所有处理程序已经在运行")
            return

        # ===== 启动前预检：数据库结构升级 + 历史邮件同步 =====
        # 目的：把“已手动发送过的舱单邮件”同步进数据库，避免自动回复时重复发送。
        try:
            import subprocess
            # 1) 数据库结构升级（补齐 sync_source 列）
            try:
                subprocess.run([sys.executable, 'UpdateDatabaseSchema.py'], check=False)
            except Exception as e:
                logger.warning(f"⚠️ 数据库结构升级脚本执行失败: {e}")


            # 修改为正确的代码：
            # 2) 同步历史邮件（关键步骤）
            from HistoryMailSync import HistoryMailSync
            sync_mgr = HistoryMailSync()
            try:
                sync_res = sync_mgr.sync_all_folders(max_emails=100, progress_callback=None)
                if sync_res.get('status') != 'completed':
                    logger.error(f"❌ 启动前历史邮件同步失败: {sync_res.get('message','未知错误')}")
                    logger.error("为避免重复发送，已阻止启动。请先修复同步问题或手动同步后再启动。")
                    return
                logger.info("✅ 启动前历史邮件同步完成")
            except Exception as e:
                logger.error(f"❌ 启动前历史邮件同步过程异常: {e}")
                logger.error("为避免重复发送，已阻止启动。请先修复同步问题或手动同步后再启动。")
                return
            logger.info("✅ 启动前历史邮件同步完成")
        except Exception as e:
            logger.error(f"❌ 启动前预检失败: {e}")
            logger.error("为避免重复发送，已阻止启动。")
            return
            
        logger.info("🚀 启动舱单邮件处理系统...")
        logger.info("=" * 60)
        logger.info("📧 系统配置:")
        logger.info("   - 进口舱单处理程序: 运行中")
        logger.info("   - 出口舱单处理程序: 运行中")
        logger.info("   - 日志文件: 分开记录")
        logger.info("   - 数据库: 分开存储")
        logger.info("=" * 60)
        
        self.running = True
        
        # 启动进口舱单处理程序
        self.import_processor.start()
        
        # 稍微延迟一下，避免同时启动造成资源竞争
        time.sleep(2)
        
        # 启动出口舱单处理程序
        self.export_processor.start()
        
        logger.info("✅ 所有处理程序已启动完成")
        logger.info("📊 系统运行中，按 Ctrl+C 停止...")
        
        # 保持主线程运行
        try:
            while self.running:
                time.sleep(1)
                # 检查处理器状态
                if not self.import_processor.thread.is_alive():
                    logger.warning("⚠️ 进口舱单处理程序已停止，尝试重启...")
                    self.import_processor.stop()
                    time.sleep(5)
                    self.import_processor.start()
                    
                if not self.export_processor.thread.is_alive():
                    logger.warning("⚠️ 出口舱单处理程序已停止，尝试重启...")
                    self.export_processor.stop()
                    time.sleep(5)
                    self.export_processor.start()
                    
        except KeyboardInterrupt:
            self.stop_all()
            
    def stop_all(self):
        """停止所有处理程序"""
        if not self.running:
            return
            
        logger.info("🛑 正在停止所有处理程序...")
        self.running = False
        
        # 发送手动关闭通知
        logger.info("📱 正在发送手动关闭短信通知...")
        
        # 发送进口舱单手动关闭通知
        import_sms_result = self.import_processor.send_manual_stop_notification()
        
        # 发送出口舱单手动关闭通知
        export_sms_result = self.export_processor.send_manual_stop_notification()
        
        # 等待短信发送完成
        time.sleep(2)
        
        # 停止进口舱单处理程序
        self.import_processor.stop()
        
        # 停止出口舱单处理程序
        self.export_processor.stop()
        
        # 总结短信发送结果
        print("\n" + "=" * 60)
        print("📱 短信通知发送结果:")
        print(f"   进口舱单: {'✅ 成功' if import_sms_result else '❌ 失败'}")
        print(f"   出口舱单: {'✅ 成功' if export_sms_result else '❌ 失败'}")
        print("=" * 60)
        
        logger.info("👋 所有处理程序已停止")
        
    def view_status(self):
        """查看系统状态"""
        print("=" * 60)
        print("📊 舱单邮件处理系统状态")
        print("=" * 60)
        print(f"进口舱单处理程序: {'✅ 运行中' if self.import_processor.running else '❌ 已停止'}")
        print(f"出口舱单处理程序: {'✅ 运行中' if self.export_processor.running else '❌ 已停止'}")
        print("=" * 60)
        
        if not self.running:
            print("使用 'start' 命令启动系统")
            print("使用 'import view' 查看进口舱单数据库")
            print("使用 'export view' 查看出口舱单数据库")
            print("使用 'import log' 查看进口舱单日志")
            print("使用 'export log' 查看出口舱单日志")
        print("=" * 60)

def handle_command_line():
    """处理命令行参数"""
    if len(sys.argv) > 1:
        command = sys.argv[1]
        
        if command == 'start':
            # 立即输出启动信息
            print("=" * 60)
            print("🚀 正在启动舱单邮件处理系统...")
            print("📧 系统配置:")
            print("   - 进口舱单处理程序: 启动中")
            print("   - 出口舱单处理程序: 启动中")
            print("   - 日志文件: 分开记录")
            print("   - 数据库: 分开存储")
            print("=" * 60)
            sys.stdout.flush()  # 刷新输出缓冲区
            
            # 启动系统
            controller = MainController()
            
            # 设置信号处理器
            def signal_handler(sig, frame):
                print("\n🛑 收到停止信号")
                controller.stop_all()
                sys.exit(0)
                
            signal.signal(signal.SIGINT, signal_handler)
            signal.signal(signal.SIGTERM, signal_handler)
            
            controller.start_all()
            
        elif command == 'status':
            controller = MainController()
            controller.view_status()
            
        elif command == 'import':
            if len(sys.argv) > 2:
                subcommand = sys.argv[2]
                if subcommand == 'view':
                    from InputAutoRW_FullFunc_2_0 import view_database_simple
                    view_database_simple()
                elif subcommand == 'log':
                    from InputAutoRW_FullFunc_2_0 import view_log_summary
                    view_log_summary()
                elif subcommand == 'test_sms':
                    from InputAutoRW_FullFunc_2_0 import send_exit_notification
                    print("🧪 测试进口舱单短信通知...")
                    print("正在发送测试短信...")
                    # 测试异常退出通知
                    send_exit_notification("这是测试异常信息", is_manual=False)
                    print("已发送异常退出测试短信")
                    
                    # 测试手动关闭通知
                    send_exit_notification(is_manual=True)
                    print("已发送手动关闭测试短信")
                    
                    print("✅ 测试完成，请检查手机是否收到短信")
                else:
                    print_usage()
            else:
                print("进口舱单处理程序命令:")
                print("  python AutoRW_MainController_fixed.py import view      # 查看数据库")
                print("  python AutoRW_MainController_fixed.py import log       # 查看处理日志")
                print("  python AutoRW_MainController_fixed.py import test_sms  # 测试短信通知")
                
        elif command == 'export':
            if len(sys.argv) > 2:
                subcommand = sys.argv[2]
                if subcommand == 'view':
                    from OutputAutoRWwithSend_3_0 import view_database_simple
                    view_database_simple()
                elif subcommand == 'log':
                    from OutputAutoRWwithSend_3_0 import view_log_summary
                    view_log_summary()
                elif subcommand == 'test_sms':
                    from OutputAutoRWwithSend_3_0 import send_exit_notification
                    print("🧪 测试出口舱单短信通知...")
                    print("正在发送测试短信...")
                    # 测试异常退出通知
                    send_exit_notification("这是测试异常信息", is_manual=False)
                    print("已发送异常退出测试短信")
                    
                    # 测试手动关闭通知
                    send_exit_notification(is_manual=True)
                    print("已发送手动关闭测试短信")
                    
                    print("✅ 测试完成，请检查手机是否收到短信")
                else:
                    print_usage()
            else:
                print("出口舱单处理程序命令:")
                print("  python AutoRW_MainController_fixed.py export view      # 查看数据库")
                print("  python AutoRW_MainController_fixed.py export log       # 查看处理日志")
                print("  python AutoRW_MainController_fixed.py export test_sms  # 测试短信通知")
                
        else:
            print_usage()
    else:
        print_usage()

def print_usage():
    """打印使用说明"""
    print("=" * 60)
    print("📦 舱单邮件自动处理系统 - 主控制器 (修复版)")
    print("=" * 60)
    print("使用方法:")
    print("  python AutoRW_MainController_fixed.py start         # 启动所有处理程序")
    print("  python AutoRW_MainController_fixed.py status        # 查看系统状态")
    print("")
    print("进口舱单处理:")
    print("  python AutoRW_MainController_fixed.py import view      # 查看数据库")
    print("  python AutoRW_MainController_fixed.py import log       # 查看处理日志")
    print("  python AutoRW_MainController_fixed.py import test_sms  # 测试短信通知")
    print("")
    print("出口舱单处理:")
    print("  python AutoRW_MainController_fixed.py export view      # 查看数据库")
    print("  python AutoRW_MainController_fixed.py export log       # 查看处理日志")
    print("  python AutoRW_MainController_fixed.py export test_sms  # 测试短信通知")
    print("")
    print("📝 说明:")
    print("  - 进口和出口舱单处理程序会同时运行")
    print("  - 每个处理程序有自己的数据库和日志文件")
    print("  - 系统会自动监控处理程序状态，异常退出时会重启")
    print("  - 手动关闭时会发送短信通知，并显示发送结果")
    print("=" * 60)

if __name__ == "__main__":
    try:
        handle_command_line()
    except KeyboardInterrupt:
        print("\n👋 用户中断操作")
    except Exception as e:
        logger.error(f"❌ 系统异常: {e}")
        traceback.print_exc()
        sys.exit(1)
