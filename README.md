舱单邮件自动处理系统
---
欢迎使用**舱单邮件自动处理系统**！本系统是一个自动处理进口和出口舱单邮件的工具。
它能够自动检查邮箱中的新邮件，提取附件中的舱单文件（TXT格式），解析其中的关键信息（如箱号、货名、提单号），并根据预设的关键词（如Calcium Nitrate, Magnesium Nitrate等）筛选出相关记录，生成Excel文件作为邮件附件回复给发件人。系统还具备日志记录、数据库存储和短信通知功能。

---
#主要功能

    自动邮件检查：定期检查指定邮箱的新邮件。
    附件解析：自动解析邮件中的TXT附件，识别进口或出口舱单格式。
    关键词筛选：根据预设的关键词筛选出相关的舱单记录。
    Excel生成：将筛选后的数据生成Excel文件，包含提单号、箱号、货名三列。
    自动回复：将生成的Excel文件作为附件自动回复给发件人。
    日志记录：记录所有邮件的处理状态，自动清理3天前的日志。
    数据库存储：将匹配到关键词并已发送Excel的邮件信息保存到SQLite数据库，保留90天记录。
    短信通知：在程序异常退出或手动关闭时发送短信通知。
    主控制器：同时管理进口和出口舱单处理程序，监控其运行状态，异常时自动重启。
---
#系统要求

    win10/11+Python 3.1或更高版本
---
#环境安装

    如果您还没有安装python，请根据以下步骤完成：
        第一步：安装Python
            访问Python官网：https://www.python.org/downloads/
            下载Python 3.1或更高版本 
            运行安装程序，务必勾选"Add Python to PATH"
            完成安装
        第二步：打开命令提示符进行以下操作
            查看版本python --version
            进入项目文件夹cd /d D:\project（根据实际进行修改）
            安装必要的包pip install openpyxl
---
#项目结构

    信息不自动读取邮箱文件2.0合并版/
        ├── AutoRW_GUI.py               # 新增：图形界面主程序
        ├── config.ini                  # 新增：配置文件
        ├── requirements.txt            # 新增：依赖包列表
    	├── AutoRW_MainController.py     # 主控制器脚本
    	├── InputAutoRW_FullFunc_2_0.py  # 进口舱单处理程序
    	├── OutputAutoRWwithSend_3_0.py  # 出口舱单处理程序
    	├── processed_emails_import.db   # 进口舱单数据库（自动生成）
    	├── processed_emails.db          # 出口舱单数据库（自动生成）
    	├── email_processing_log_import.csv  # 进口舱单日志（自动生成）
    	└── email_processing_log.csv     # 出口舱单日志（自动生成）
---
#运行项目命令行

    1.启动所有处理程序
        python AutoRW_MainController.py start
    2.查看系统状态
        python AutoRW_MainController.py status
    3.查看进口舱单数据库
        python AutoRW_MainController.py import view
    4.查看出口舱单数据库
        python AutoRW_MainController.py export view
    5.查看进口舱单日志
        python AutoRW_MainController.py import log
    6.查看出口舱单日志
        python AutoRW_MainController.py export log
    7.测试短信通知
        python AutoRW_MainController.py import test_sms
        python AutoRW_MainController.py export test_sms
---
#修改参数

    首次使用，需要修改代码 InputAutoRW_FullFunc_2_0.py 中 27-50 行
        email_address = "xxx@coscoshipping.com"  # 替换为您的邮箱
        password = "xxxxxxx"  # 替换为您的授权码（需在邮箱系统中获取）
        keywords #修改为需要的关键词
    
        SMS_ACCOUNT = "jksc034"               # 改为实际账户名
        SMS_PASSWORD = "jksc034"              # 改为实际短信发送密码
        SMS_MOBILES = "xxxxxxxx"               # 目标手机号码，多个用半角","分隔
        SMS_CONTENT_TEMPLATE = "【天津港集装箱码头有限公司】进口舱单邮件处理程序异常退出，请检查！"                                          #修改为您需要的短信发送内容
        SMS_API_URL = "http://sh2.ipyy.com/sms.aspx"   #目前为普通短信发送api
    
    同时，需要修改代码 OutputAutoRWwithSend_3_0.py 中 27-50 行相同位置的相应内容。
---
#代码逻辑

    进口舱单处理逻辑
        识别以00:IFCSUM:开头的TXT文件
        从12行提取提单号（第二个冒号分隔的字段）
        从44行和47行提取货物描述
        从51行提取箱号
        检查货物描述是否包含关键词
        生成Excel文件并发送回复邮件
    出口舱单处理逻辑
        识别以00NCLCONTAINER LIST开头的TXT文件
        从51行提取箱号（位置3-13）和提单号（位置29-44）
        从53行提取货物描述（位置13-43）
        检查货物描述是否包含关键词
        生成Excel文件并发送回复邮件
---
#日志说明
    
    一、日志文件
        email_processing_log_import.csv：进口舱单处理日志
        email_processing_log.csv：出口舱单处理日志
    二、日志字段
        timestamp：处理时间
        email_uid：邮件唯一标识
        sender：发件人
        subject：邮件主题
        has_keyword：是否匹配关键词
        excel_sent：是否已发送Excel
        matched_keywords：匹配的关键词
        container_count：匹配的箱号数量
    三、自动清理
        日志自动保留3天内的记录+数据库自动保留90天内的记录
---
#短信通知


    触发条件
        程序异常退出（自动发送）
        用户手动关闭（按Ctrl+C，自动发送）
        邮箱登录失败（自动发送）

#注意事项

    首次运行：请先配置邮箱信息
    测试运行：建议先用测试命令验证功能
    定期检查：定期查看日志了解系统运行状态
    数据备份：重要的数据库文件建议定期备份
    网络要求：需要稳定的网络连接

#成功运行示例

    🚀 启动舱单邮件自动处理系统...
    📧 系统配置：
        - 进口舱单处理程序: 运行中
        - 出口舱单处理程序: 运行中
        - 日志文件: 分开记录
        - 数据库: 分开存储
    ✅ 所有处理程序已启动完成
    📊 系统运行中，按 Ctrl+C 停止...

#成功停止示例

    🛑 收到停止信号
    🛑 正在停止所有处理程序...
    📱 正在发送手动关闭短信通知...
    📱 [进口] ✅ 短信发送成功: ...
    ✅ 进口舱单手动关闭短信通知发送成功
    📱 [出口] ✅ 短信发送成功: ...
    ✅ 出口舱单手动关闭短信通知发送成功
    ============================================================
    📱 短信通知发送结果:
        进口舱单: ✅ 成功
        出口舱单: ✅ 成功
    ============================================================
    👋 所有处理程序已停止
---
#

	版本：1.0.0
	最后更新：2025年12月
	开发语言：VSCode + Python 3.13.7
	作者：张佩盈