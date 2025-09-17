import sys
import os
import json
import random

# Windows系统相关导入
import winreg
import win32event
import win32api

# PyQt相关导入
from PyQt6.QtWidgets import QApplication, QWidget, QLabel
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QPainter, QPainterPath, QRegion, QGuiApplication

# 配置常量类
class Config:
    WINDOW_WIDTH = 500
    WINDOW_HEIGHT = 45
    SCROLL_SPEED = 1
    QUOTE_UPDATE_INTERVAL = 3 * 60 * 1000  # 3分钟
    VISIBILITY_CHECK_INTERVAL = 1000  # 1秒
    Z_ORDER_CHECK_INTERVAL = 500  # 0.5秒检查一次窗口层级
    MUTEX_NAME = "人民日报金句应用"
    APP_NAME = "人民日报金句"
    DEFAULT_QUOTE = "每一天都是新的开始"

class MomentumQuotesApp(QWidget):
    def __init__(self):
        super().__init__()
        
        # 初始化定时器属性
        self.scroll_timer = None
        self.timer = None
        self.visibility_timer = None
        self.z_order_timer = None
        
        # 初始化滚动相关属性
        self.scroll_position = 0
        self.is_scrolling = False
        self.text_width = 0
        
        # 初始化流程
        self._setup_window_properties()
        self._init_ui_components()
        
        # 加载数据和启动服务
        self.quotes = self.load_quotes()
        if self.quotes:
            self.show_random_quote()
        
        # 启动定时任务
        self._start_timers()
        
        # 确保窗口显示和置顶
        self.show()
        self.raise_()
    
    def closeEvent(self, event):
        """窗口关闭事件，确保定时器正确停止"""
        self._stop_all_timers()
        super().closeEvent(event)
    
    def _stop_all_timers(self):
        """停止所有定时器"""
        for timer_name in ['scroll_timer', 'timer', 'visibility_timer', 'z_order_timer']:
            timer = getattr(self, timer_name, None)
            if timer and timer.isActive():
                timer.stop()
    
    def _setup_window_properties(self):
        """设置窗口基本属性"""
        # 设置窗口标志
        self.setWindowFlags(
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowDoesNotAcceptFocus
        )
        
        # 设置透明相关属性
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        
        # 设置窗口尺寸和遮罩
        self.resize(Config.WINDOW_WIDTH, Config.WINDOW_HEIGHT)
        self.set_mask_region()
        
        # 定位窗口
        self.position_window()
    
    def _init_ui_components(self):
        """初始化UI组件"""
        # 创建标签
        self.quote_label = QLabel(self)
        self.quote_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.quote_label.setWordWrap(False)
        self.quote_label.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        self.quote_label.setStyleSheet("color: rgba(30, 30, 30, 0.85); background: transparent;")
        self.quote_label.resize(Config.WINDOW_WIDTH + 1000, Config.WINDOW_HEIGHT)
        self.quote_label.move(15, 0)
    
    def _start_timers(self):
        """启动所有定时器"""
        # 设置定时任务，每3分钟切换一条语录
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.show_random_quote)
        self.timer.start(Config.QUOTE_UPDATE_INTERVAL)
        
        # 添加窗口可见性检查定时器
        self.visibility_timer = QTimer(self)
        self.visibility_timer.timeout.connect(self.check_visibility)
        self.visibility_timer.start(Config.VISIBILITY_CHECK_INTERVAL)
        
        # 添加窗口层级检查定时器
        self.z_order_timer = QTimer(self)
        self.z_order_timer.timeout.connect(self.check_z_order)
        self.z_order_timer.start(Config.Z_ORDER_CHECK_INTERVAL)
    
    def set_mask_region(self):
        """设置窗口遮罩区域"""
        path = QPainterPath()
        path.addRect(0, 0, self.width(), self.height())
        mask = QRegion(path.toFillPolygon().toPolygon())
        self.setMask(mask)
    
    def position_window(self):
        """将窗口定位在任务栏上方，支持多显示器"""
        try:
            # 获取当前窗口所在的屏幕或主屏幕
            screen = self.screen() or QGuiApplication.primaryScreen()
            screen_geometry = screen.geometry()
            
            # 计算窗口位置（左下角）
            x_position = screen_geometry.x() + 10
            bottom_offset = 50 if screen_geometry.height() > 1080 else 45
            y_position = screen_geometry.y() + screen_geometry.height() - bottom_offset
            
            self.move(x_position, y_position)
        except Exception as e:
            print(f"多显示器定位出错: {e}")
            # 回退到简单的主屏幕定位
            screen = QGuiApplication.primaryScreen().geometry()
            x_position = 10
            y_position = screen.height() - 50 if screen.height() > 1080 else screen.height() - 45
            self.move(x_position, y_position)
    
    def show_random_quote(self):
        """显示一条随机的励志语录"""
        # 随机选择并显示语录
        quote = random.choice(self.quotes) if self.quotes else {}
        full_text = quote.get('quote', Config.DEFAULT_QUOTE)
        
        # 设置字体样式并更新标签内容
        self._set_label_style()
        self.quote_label.setText(full_text)
        
        # 重置窗口和标签属性
        self.resize(Config.WINDOW_WIDTH, Config.WINDOW_HEIGHT)
        self.set_mask_region()
        self.quote_label.resize(Config.WINDOW_WIDTH + 1000, Config.WINDOW_HEIGHT)
        self.quote_label.move(15, 0)
        
        # 智能判断是否需要滚动
        self.check_scroll_needed()
        self.raise_()  # 更换语录后确保窗口置顶
    
    def _set_label_style(self):
        """设置标签字体样式"""
        # 创建字体，添加字体回退列表
        font = QFont("Segoe UI Variable, 微软雅黑, Microsoft YaHei, 宋体, SimSun", 13, QFont.Weight.Light)
        font.setStyleStrategy(QFont.StyleStrategy.PreferAntialias)
        self.quote_label.setFont(font)
    
    def check_scroll_needed(self):
        """检查文本是否需要滚动"""
        # 重置滚动属性
        self.scroll_position = 0
        self.is_scrolling = False
        
        # 计算文本宽度和可用宽度
        self.text_width = self.quote_label.fontMetrics().horizontalAdvance(self.quote_label.text())
        available_width = self.width() - 30
        
        # 判断是否需要滚动
        if self.text_width > available_width:
            if not self.scroll_timer:
                self.scroll_timer = QTimer(self)
                self.scroll_timer.timeout.connect(self.scroll_text)
            elif self.scroll_timer.isActive():
                self.scroll_timer.stop()
            
            # 开始滚动
            self.is_scrolling = True
            self.scroll_timer.start(50)
        elif self.scroll_timer and self.scroll_timer.isActive():
            self.scroll_timer.stop()
            self.quote_label.move(15, 0)
    
    def scroll_text(self):
        """实现平滑循环滚动效果"""
        if not self.is_scrolling:
            return
        
        # 缓存计算结果
        available_width = self.width() - 30
        max_scroll = max(0, self.text_width - available_width)
        
        # 更新滚动位置
        self.scroll_position = (self.scroll_position + Config.SCROLL_SPEED) % (max_scroll + available_width * 2)
        
        # 计算新位置
        if self.scroll_position <= max_scroll:
            new_pos = 15 - self.scroll_position
        else:
            buffer_position = self.scroll_position - max_scroll
            new_pos = 15 - max_scroll - available_width + buffer_position
        
        # 只在位置变化时才移动标签
        if new_pos != self.quote_label.pos().x():
            self.quote_label.move(new_pos, 0)
    
    def check_visibility(self):
        """检查窗口可见性并在需要时恢复显示"""
        if not self.isVisible():
            self.show()
            self.raise_()
    
    def check_z_order(self):
        """定期检查窗口层级，确保始终置顶"""
        self.raise_()
    
    def showEvent(self, event):
        """窗口显示事件"""
        super().showEvent(event)
        self.raise_()
    
    def resizeEvent(self, event):
        """窗口大小改变事件"""
        super().resizeEvent(event)
        self.set_mask_region()
        if self.quotes:
            self.check_scroll_needed()
    
    # 忽略所有事件，实现操作透明
    def keyPressEvent(self, event):
        event.ignore()
    
    def keyReleaseEvent(self, event):
        event.ignore()
    
    def mousePressEvent(self, event):
        event.ignore()
    
    def mouseReleaseEvent(self, event):
        event.ignore()
    
    def mouseMoveEvent(self, event):
        event.ignore()
    
    def wheelEvent(self, event):
        event.ignore()
    
    def load_quotes(self):
        """加载励志语录数据，包含错误处理和数据验证"""
        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "quotes.json")
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 数据验证
                if isinstance(data, list):
                    validated_quotes = [
                        {'quote': str(item['quote'])} 
                        for item in data 
                        if isinstance(item, dict) and 'quote' in item and item['quote']
                    ]
                    # 返回有效数据或默认值
                    return validated_quotes if validated_quotes else [{'quote': Config.DEFAULT_QUOTE}]
                return [{'quote': Config.DEFAULT_QUOTE}]
        except (json.JSONDecodeError, FileNotFoundError, Exception) as e:
            error_type = (
                "JSON文件格式错误" if isinstance(e, json.JSONDecodeError) else
                "文件不存在" if isinstance(e, FileNotFoundError) else
                f"加载数据出错: {e}"
            )
            print(error_type)
            return [{'quote': Config.DEFAULT_QUOTE}]
    
    def paintEvent(self, event):
        """自定义绘制，移除背景"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

def add_to_startup():
    """添加应用到开机自启动"""
    key_path = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    try:
        # 检查是否已添加
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ) as key:
            try:
                winreg.QueryValueEx(key, Config.APP_NAME)
                return  # 已存在，不需要再次添加
            except FileNotFoundError:
                pass
        
        # 获取当前可执行文件路径
        exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        
        # 添加到开机自启动
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, Config.APP_NAME, 0, winreg.REG_SZ, exe_path)
    except Exception as e:
        print(f"添加开机自启动失败: {e}")

if __name__ == "__main__":
    # 实现单例模式
    try:
        mutex = win32event.CreateMutex(None, True, Config.MUTEX_NAME)
        if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            sys.exit(0)  # 已存在实例，退出当前程序
    except Exception as e:
        print(f"创建互斥体失败: {e}")
        sys.exit(1)
    
    # 添加到开机自启动
    add_to_startup()
    
    # 启动应用
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MomentumQuotesApp()
    sys.exit(app.exec())