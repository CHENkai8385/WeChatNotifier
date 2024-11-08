import uiautomation as uia
import re
from win11toast import notify
import time
from datetime import datetime
import pystray
from PIL import Image
import threading
import os

class WeChat():
    SessionItemList: list = []

    def __init__(
            self,
        ) -> None:
        self.UiaAPI: uia.WindowControl = uia.WindowControl(ClassName='WeChatMainWndForPC', searchDepth=1)
        MainControl1 = [i for i in self.UiaAPI.GetChildren() if not i.ClassName][0]
        MainControl2 = MainControl1.GetFirstChildControl()
        self.NavigationBox, self.SessionBox, self.ChatBox  = MainControl2.GetChildren()
        
        # 初始化导航栏，以A开头 | self.NavigationBox  -->  A_xxx
        self.A_MyIcon = self.NavigationBox.ButtonControl()
        
        self.nickname = self.A_MyIcon.Name
        print(f'初始化成功，获取到已登录窗口：{self.nickname}')

        self.message_cache = {}  # 用于存储已发送的消息
        self.cache_expire_time = 3600  # 缓存过期时间（秒）

    def print_all_children(self, control, level=0):
        """打印控件的所有子级，包含详细控件信息
        
        Args:
            control: 要打印子级控件
            level (int): 当前递归层级，用于缩进显示
        """
        control_info = [
            # f'ClassName: {control.ClassName}',
            f'Name: {control.Name}',
            # f'AutomationId: {control.AutomationId}',
            # f'ControlType: {control.ControlType}',
            f'ControlTypeName: {control.ControlTypeName}',
            # f'BoundingRectangle: {control.BoundingRectangle}',
            # f'IsEnabled: {control.IsEnabled}',
            # f'IsOffscreen: {control.IsOffscreen}'
        ]
        print(' ' * level+str(level)+ '- Control Properties:')
        for info in control_info:
            print(' ' * (level + 1) + info)
        
        for child in control.GetChildren():
            self.print_all_children(child, level + 1)

            
    def has_new_message(self, SessionItem) -> bool:
        """检查会话项是否包含新消息
        
        Args:
            SessionItem: 会话项控件
            
        Returns:
            bool: 如果包含新消息返回True，否则返回False
        """
        return bool(re.search('\d+条新消息', SessionItem.Name))

    def GetNewSessionInfo(self, SessionItem):
        """获取会话的新消息信息
        
        Args:
            SessionItem: 会话项控件
            
        Returns:
            dict: 包含会话信息的字典，获取失败则返回None
        """
        try:
            # 获取新消息数量
            amount = int([i for i in SessionItem.GetFirstChildControl().GetChildren() 
                         if type(i) == uia.TextControl][0].Name)
            children = SessionItem.PaneControl().PaneControl().PaneControl().GetChildren()
            time = children[2].Name if children[2].Name else children[3].Name
            msg = SessionItem.PaneControl().PaneControl().GetChildren()[1].TextControl().Name
            sessionname = SessionItem.PaneControl().ButtonControl().Name
            
            info = {
                "name": sessionname,
                "amount": amount,
                "time": time,
                "msg": msg
            }
            return info
        except Exception as e:
            print(f"获取会话信息失败: {str(e)}")
            return None

    def GetSessionList(self):
        self.SessionItem = self.SessionBox.ListItemControl()
        SessionList = {}
        while self.SessionItem is not None:
            if self.has_new_message(self.SessionItem):
                info = self.GetNewSessionInfo(self.SessionItem)
                if info['name'] not in SessionList:
                    SessionList[info['name']] = info
            self.SessionItem = self.SessionItem.GetNextSiblingControl()
        return SessionList

    def _generate_message_key(self, info):
        """生成消息的唯一标识
        
        Args:
            info (dict): 消息信息字典
            
        Returns:
            str: 消息的唯一标识
        """
        # 提取发送者名称（如果消息中包含）
        sender = ""
        msg = info["msg"]
        if "：" in msg:
            sender = msg.split("：")[0]
            msg = msg.split("：")[1]
        
        # 组合唯一标识：会话名称 + 发送者 + 消息内容 + 时间
        return f"{info['name']}_{sender}_{msg}_{info['time']}"

    def _should_send_notification(self, message_key, info):
        """检查是否应该发送通知
        
        Args:
            message_key (str): 消息唯一标识
            info (dict): 消息信息字典
            
        Returns:
            bool: 是否应该发送通知
        """
        current_time = time.time()
        
        # # 清理过期缓存
        # expired_keys = [k for k, v in self.message_cache.items() 
        #                if current_time - v > self.cache_expire_time]
        # for k in expired_keys:
        #     del self.message_cache[k]
        
        # 检查是否是新消息
        if message_key not in self.message_cache:
            self.message_cache[message_key] = current_time
            return True
        return False

    def _format_notification_content(self, info):
        """格式化通知内容
        
        Args:
            info (dict): 消息信息字典
            
        Returns:
            tuple: (标题, 内容)
        """
        # 提取发送者信息
        sender = "未知发送者"
        content = info["msg"]
        # info['time']
        title = f"{info['name']}（{info['amount']}）"
        
        body = f"{content}"
        
        return title, body

    def send_notifications(self, session_list):
        """发送系统通知
        
        Args:
            session_list (dict): 会话列表
        """
        for session_name, info in session_list.items():
            message_key = self._generate_message_key(info)
            
            if self._should_send_notification(message_key, info):
                title, body = self._format_notification_content(info)
                try:
                    notify(
                        title=title,
                        body=body,
                        app_id='微信消息通知',
                        audio={'silent': True},
                        duration="short"  # 或 "long"
                    )
                    print(f"已发送通知: {title}")
                except Exception as e:
                    print(f"发送通知失败: {str(e)}")

    def start_monitoring(self):
        """开始监控微信消息"""
        while not self.should_exit:
            try:
                session_list = self.GetSessionList()
                self.send_notifications(session_list)
                time.sleep(5)
            except Exception as e:
                print(f"监控过程出错: {str(e)}")
                time.sleep(5)

def create_icon():
    # 使用自定义图标文件
    icon_path = os.path.join(os.path.dirname(__file__), "assets/images/wechat.png")
    return Image.open(icon_path)

def setup_tray():
    def on_exit(icon, item):
        icon.stop()
        wx.should_exit = True

    # 创建系统托盘图标
    icon = pystray.Icon(
        "wechat_notifier",
        create_icon(),
        "微信消息提醒",
        menu=pystray.Menu(
            pystray.MenuItem("退出", on_exit)
        )
    )
    return icon

if __name__ == '__main__':
    wx = WeChat()
    wx.should_exit = False

    # 创建并启动系统托盘图标
    icon = setup_tray()
    
    # 在单独的线程中运行消息监控
    monitor_thread = threading.Thread(target=wx.start_monitoring, daemon=True)
    monitor_thread.start()
    
    # 运行系统托盘图标（这会阻塞主线程）
    icon.run()