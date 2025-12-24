"""
Windows表情包管理器
功能：通过全局热键快速呼出，实现表情包分类管理、快速检索和便捷使用
"""

import sys
import os
import json
import logging
import threading
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime
from collections import OrderedDict

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QScrollArea, QPushButton, QLineEdit, QLabel, QSystemTrayIcon,
    QMenu, QFileDialog, QMessageBox, QGridLayout, QFrame
)
from PySide6.QtCore import Qt, QSize, QTimer, Signal, QThread, QObject
from PySide6.QtGui import (
    QPixmap, QImage, QIcon, QPainter, QColor, QPalette,
    QClipboard, QKeySequence, QShortcut
)
from PIL import Image
import keyboard


# ==================== 资源路径处理 ====================
def resource_path(relative_path):
    """获取资源的绝对路径。打包到PyInstaller后，使用临时文件夹路径"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# 图标路径
try:
    program_icon = resource_path("assets/st.ico")
    # 检查文件是否存在
    if not os.path.exists(program_icon):
        # 尝试相对路径
        base_dir = os.path.dirname(os.path.abspath(__file__))
        program_icon = os.path.join(base_dir, "assets", "st.ico")
        if not os.path.exists(program_icon):
            # 如果都没有，创建一个临时的图标文件
            program_icon = None
            logging.warning("图标文件未找到，将使用默认图标")
except Exception as e:
    program_icon = None
    logging.warning(f"图标路径处理失败: {e}")

logging.info(f"图标路径: {program_icon}")

# ==================== 配置管理器 ====================
class ConfigManager:
    """JSON配置文件管理"""

    def __init__(self):
        self.config_dir = Path.cwd() / ".sticker_manager"
        self.config_file = self.config_dir / "config.json"
        self.config = self.load_config()

    def load_config(self) -> dict:
        """加载配置文件"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logging.error(f"配置文件加载失败: {e}")

        # 返回默认配置
        return self.get_default_config()

    def get_default_config(self) -> dict:
        """获取默认配置"""
        return {
            "library_path": "",
            "hotkey": "ctrl+shift+e",
            "window_position": [900, 50],
            "window_size": [600, 430],
            "ui": {
                "category_button_size": 90,
                "grid_cell_size": 120,
                "grid_columns": 3,
            },
            "behavior": {
                "copy_on_double_click": True,
                "highlight_on_click": True,
                "search_delay_ms": 300,
            },
            "performance": {"thumbnail_cache_size": 200, "lazy_load_enabled": True},
        }

    def save_config(self):
        """保存配置文件"""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            logging.info("配置已保存")
        except Exception as e:
            logging.error(f"配置保存失败: {e}")

    def get(self, key: str, default=None):
        """获取配置项"""
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict):
                value = value.get(k, default)
            else:
                return default
        return value if value is not None else default

    def set(self, key: str, value):
        """设置配置项"""
        keys = key.split('.')
        config = self.config
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        config[keys[-1]] = value


# ==================== 日志系统 ====================
def setup_logging():
    """配置日志系统"""
    log_dir = Path.home() / ".sticker_manager" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / f"sticker_manager_{datetime.now().strftime('%Y%m%d')}.log"

    logging.basicConfig(
        level=logging.INFO,
        format='[%(asctime)s] [%(levelname)s] [%(name)s] - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )


# ==================== 数据管理器 ====================
class StickerLibrary:
    """表情库数据管理"""

    SUPPORTED_FORMATS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}

    def __init__(self, library_path: str):
        self.library_path = Path(library_path)
        self.categories: Dict[str, List[Path]] = {}
        self.all_stickers: List[Path] = []

    def load_library(self):
        """扫描并加载表情库结构"""
        if not self.library_path.exists():
            logging.warning(f"表情库路径不存在: {self.library_path}")
            return

        self.categories.clear()
        self.all_stickers.clear()

        logging.info(f"开始扫描表情库: {self.library_path}")

        # 遍历子目录作为分类
        for category_dir in self.library_path.iterdir():
            if not category_dir.is_dir():
                continue

            category_name = category_dir.name
            stickers = []

            # 扫描该分类下的所有图片文件
            for file_path in category_dir.iterdir():
                if file_path.is_file() and file_path.suffix.lower() in self.SUPPORTED_FORMATS:
                    stickers.append(file_path)
                    self.all_stickers.append(file_path)

            if stickers:
                self.categories[category_name] = sorted(stickers, key=lambda x: x.name)
                logging.info(f"分类 [{category_name}] 加载 {len(stickers)} 个表情")

        logging.info(f"表情库加载完成，共 {len(self.categories)} 个分类，{len(self.all_stickers)} 个表情")

    def search_stickers(self, keyword: str) -> List[Path]:
        """搜索表情（按文件名）"""
        if not keyword:
            return []

        keyword = keyword.lower()
        results = []

        for sticker in self.all_stickers:
            if keyword in sticker.stem.lower():
                results.append(sticker)

        return results


# ==================== 图像缓存管理器 ====================
class ThumbnailCache:
    """LRU缩略图缓存"""

    def __init__(self, max_size: int = 200):
        self.cache: OrderedDict[str, QPixmap] = OrderedDict()
        self.max_size = max_size

    def get(self, key: str) -> Optional[QPixmap]:
        """获取缓存的缩略图"""
        if key in self.cache:
            # 移到最后（最近使用）
            self.cache.move_to_end(key)
            return self.cache[key]
        return None

    def put(self, key: str, pixmap: QPixmap):
        """添加缩略图到缓存"""
        if key in self.cache:
            self.cache.move_to_end(key)
        else:
            self.cache[key] = pixmap
            # 超出容量则删除最旧的
            if len(self.cache) > self.max_size:
                self.cache.popitem(last=False)

    def clear(self):
        """清空缓存"""
        self.cache.clear()


# ==================== UI组件 ====================
class StickerCell(QFrame):
    """表情单元格组件"""

    clicked = Signal(Path)
    double_clicked = Signal(Path)

    def __init__(self, sticker_path: Path, cell_size: int, parent=None):
        super().__init__(parent)
        self.sticker_path = sticker_path
        self.cell_size = cell_size
        self.is_highlighted = False

        self.setFixedSize(cell_size, cell_size)
        self.setFrameShape(QFrame.Box)
        self.setStyleSheet("QFrame { background-color: #f5f5f5; border: 2px solid #e0e0e0; }")

        # 图片标签
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setScaledContents(False)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.addWidget(self.image_label)

    def set_thumbnail(self, pixmap: QPixmap):
        """设置缩略图"""
        if pixmap:
            # 等比例缩放
            scaled_pixmap = pixmap.scaled(
                self.cell_size - 10,
                self.cell_size - 10,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.image_label.setPixmap(scaled_pixmap)

    def mousePressEvent(self, event):
        """单击高亮"""
        self.is_highlighted = True
        self.setStyleSheet("QFrame { background-color: #e3f2fd; border: 2px solid #2196f3; }")
        self.clicked.emit(self.sticker_path)

    def mouseDoubleClickEvent(self, event):
        """双击复制"""
        self.double_clicked.emit(self.sticker_path)

    def clear_highlight(self):
        """清除高亮"""
        self.is_highlighted = False
        self.setStyleSheet("QFrame { background-color: #f5f5f5; border: 2px solid #e0e0e0; }")


class CategoryButton(QPushButton):
    """分类按钮组件"""

    def __init__(self, category_name: str, first_sticker: Path, button_size: int, parent=None):
        super().__init__(parent)
        self.category_name = category_name
        self.first_sticker = first_sticker

        self.setFixedSize(button_size, button_size)
        self.setToolTip(category_name)
        self.setStyleSheet("""
            QPushButton {
                background-color: #ffffff;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
            }
            QPushButton:hover {
                border: 2px solid #2196f3;
                background-color: #e3f2fd;
            }
            QPushButton:pressed {
                background-color: #bbdefb;
            }
        """)

    def set_thumbnail(self, pixmap: QPixmap):
        """设置缩略图为按钮图标"""
        if pixmap:
            scaled_pixmap = pixmap.scaled(
                self.width() - 10,
                self.height() - 10,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.setIcon(QIcon(scaled_pixmap))
            self.setIconSize(QSize(self.width() - 10, self.height() - 10))


# ==================== 全局热键监听 ====================
class HotkeyListener(QObject):
    """全局热键监听器"""

    hotkey_pressed = Signal()

    def __init__(self, hotkey: str):
        super().__init__()
        self.hotkey = hotkey
        self.running = True
        self.thread = None

    def start(self):
        """启动热键监听"""
        self.running = True
        try:
            keyboard.add_hotkey(self.hotkey, self.on_hotkey)
            logging.info(f"全局热键已注册: {self.hotkey}")
        except Exception as e:
            logging.error(f"热键注册失败: {e}")

    def stop(self):
        """停止热键监听"""
        self.running = False
        try:
            keyboard.remove_hotkey(self.hotkey)
        except:
            pass
        try:
            keyboard.unhook_all()
        except:
            pass
        logging.info("热键监听已停止")

    def on_hotkey(self):
        """热键触发"""
        if self.running:
            self.hotkey_pressed.emit()


def get_existing_tray_icon() -> QSystemTrayIcon | None:
    """遍历应用内所有对象，找到第一个QSystemTrayIcon实例"""
    app = QApplication.instance()
    if not app:
        return None

    # 递归查找 QSystemTrayIcon
    def find_tray_in_children(parent) -> QSystemTrayIcon | None:
        # 检查当前对象是否是 QSystemTrayIcon
        if isinstance(parent, QSystemTrayIcon):
            return parent

        # 检查所有子对象
        for child in parent.children():
            result = find_tray_in_children(child)
            if result:
                return result
        return None

    # 从应用程序实例开始查找
    return find_tray_in_children(app)


# ==================== 主窗口 ====================
class StickerManagerWindow(QMainWindow):
    """表情包管理器主窗口"""

    def __init__(self, config: ConfigManager):
        super().__init__()
        self.config = config
        self.library = None
        self.thumbnail_cache = ThumbnailCache(config.get('performance.thumbnail_cache_size', 200))
        self.current_category = None
        self.current_cells: List[StickerCell] = []

        self.init_ui()
        self.load_library()

    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("表情包管理器")

        # 设置窗口大小和位置
        size = self.config.get('window_size', [500, 700])
        pos = self.config.get('window_position', [1100, 50])
        self.setGeometry(pos[0], pos[1], size[0], size[1])

        # 不置顶
        self.setWindowFlags(Qt.Window)

        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局（水平）
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)

        # 左侧分类面板
        left_panel = self.create_category_panel()
        main_layout.addWidget(left_panel, 1)

        # 右侧表情面板
        right_panel = self.create_sticker_panel()
        main_layout.addWidget(right_panel, 3)

        # ESC键隐藏
        esc_shortcut = QShortcut(QKeySequence(Qt.Key_Escape), self)
        esc_shortcut.activated.connect(self.hide_window)

        logging.info("主窗口初始化完成")

    def create_category_panel(self) -> QWidget:
        """创建分类面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 分类滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        self.category_container = QWidget()
        self.category_layout = QVBoxLayout(self.category_container)
        self.category_layout.setAlignment(Qt.AlignTop)
        self.category_layout.setSpacing(8)

        scroll.setWidget(self.category_container)
        layout.addWidget(scroll)

        return panel

    def create_sticker_panel(self) -> QWidget:
        """创建表情面板"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 搜索框
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索表情...")
        self.search_input.textChanged.connect(self.on_search)

        clear_btn = QPushButton("✕")
        clear_btn.setFixedWidth(30)
        clear_btn.clicked.connect(lambda: self.search_input.clear())

        search_layout.addWidget(self.search_input)
        search_layout.addWidget(clear_btn)
        layout.addLayout(search_layout)

        # 表情网格滚动区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        self.grid_container = QWidget()
        self.grid_layout = QGridLayout(self.grid_container)
        self.grid_layout.setSpacing(8)
        self.grid_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        scroll.setWidget(self.grid_container)
        layout.addWidget(scroll)

        return panel

    def load_library(self):
        """加载表情库"""
        library_path = self.config.get('library_path', '')

        if not library_path or not Path(library_path).exists():
            # 弹出选择目录对话框
            library_path = QFileDialog.getExistingDirectory(
                self,
                "选择表情库目录",
                str(Path.home())
            )

            if not library_path:
                logging.warning("未选择表情库目录")
                return

            self.config.set('library_path', library_path)
            self.config.save_config()

        self.library = StickerLibrary(library_path)
        self.library.load_library()

        self.populate_categories()

    def populate_categories(self):
        """填充分类列表"""
        # 清空现有按钮
        while self.category_layout.count():
            item = self.category_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.library or not self.library.categories:
            return

        button_size = self.config.get('ui.category_button_size', 80)

        for category_name, stickers in self.library.categories.items():
            if not stickers:
                continue

            first_sticker = stickers[0]
            btn = CategoryButton(category_name, first_sticker, button_size)
            btn.clicked.connect(lambda checked, cat=category_name: self.show_category(cat))
            btn.setToolTip(category_name)

            self.category_layout.addWidget(btn)

            # 加载缩略图
            self.load_thumbnail_for_button(btn, first_sticker)

        # 默认显示第一个分类
        if self.library.categories:
            first_category = list(self.library.categories.keys())[0]
            self.show_category(first_category)

    def load_thumbnail_for_button(self, button: CategoryButton, image_path: Path):
        """为分类按钮加载缩略图"""
        pixmap = self.get_thumbnail(image_path, button.width() - 10)
        if pixmap:
            button.set_thumbnail(pixmap)

    def show_category(self, category_name: str):
        """显示指定分类的表情"""
        self.current_category = category_name
        self.search_input.clear()

        stickers = self.library.categories.get(category_name, [])
        self.display_stickers(stickers)

        logging.info(f"显示分类: {category_name}, 共 {len(stickers)} 个表情")

    def display_stickers(self, stickers: List[Path]):
        """在网格中显示表情列表"""
        # 清空现有单元格
        for cell in self.current_cells:
            cell.deleteLater()
        self.current_cells.clear()

        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not stickers:
            return

        cell_size = self.config.get('ui.grid_cell_size', 120)
        columns = self.config.get('ui.grid_columns', 3)

        # 使用独立的计数器跟踪实际添加的单元格
        idx = 0

        for _, sticker_path in enumerate(stickers):
            # 跳过预览图
            if os.path.basename(sticker_path).startswith(".preview"):
                continue

            # 使用实际添加的单元格索引计算行列
            row = idx // columns
            col = idx % columns

            cell = StickerCell(sticker_path, cell_size)
            cell.clicked.connect(self.on_sticker_clicked)
            cell.double_clicked.connect(self.on_sticker_double_clicked)
            cell.setToolTip(os.path.basename(sticker_path))

            self.grid_layout.addWidget(cell, row, col)
            self.current_cells.append(cell)

            # 加载缩略图
            self.load_thumbnail_for_cell(cell, sticker_path)

            # 递增实际单元格计数器
            idx += 1

    def load_thumbnail_for_cell(self, cell: StickerCell, image_path: Path):
        """为表情单元格加载缩略图"""
        pixmap = self.get_thumbnail(image_path, cell.cell_size - 10)
        if pixmap:
            cell.set_thumbnail(pixmap)

    def get_thumbnail(self, image_path: Path, max_size: int) -> Optional[QPixmap]:
        """获取缩略图（带缓存）"""
        cache_key = f"{image_path}_{max_size}"

        # 尝试从缓存获取
        cached = self.thumbnail_cache.get(cache_key)
        if cached:
            return cached

        # 生成缩略图
        try:
            pixmap = QPixmap(str(image_path))
            if pixmap.isNull():
                logging.warning(f"无法加载图片: {image_path}")
                return None

            # 缓存并返回
            self.thumbnail_cache.put(cache_key, pixmap)
            return pixmap

        except Exception as e:
            logging.error(f"加载缩略图失败 {image_path}: {e}")
            return None

    def on_sticker_clicked(self, sticker_path: Path):
        """表情单击事件"""
        # 清除其他单元格的高亮
        for cell in self.current_cells:
            if cell.sticker_path != sticker_path:
                cell.clear_highlight()

    def on_sticker_double_clicked(self, sticker_path: Path):
        """表情双击事件 - 复制到剪贴板"""
        try:
            pixmap = QPixmap(str(sticker_path))
            if not pixmap.isNull():
                clipboard = QApplication.clipboard()
                clipboard.setPixmap(pixmap)
                logging.info(f"已复制表情: {sticker_path.name}")

                # 可选：复制后自动隐藏窗口
                self.hide_window()
                # 获取托盘图标
                tray = get_existing_tray_icon()
                # 显示提示
                if tray:
                    tray.showMessage(
                        self.windowTitle(),
                        f"已复制表情: {sticker_path.name}",
                        msecs=1000,
                    )
        except Exception as e:
            logging.error(f"复制表情失败: {e}")

    def on_search(self, text: str):
        """搜索事件"""
        if not text.strip():
            # 恢复当前分类显示
            if self.current_category:
                self.show_category(self.current_category)
            return

        # 延迟搜索
        if hasattr(self, 'search_timer'):
            self.search_timer.stop()

        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(lambda: self.perform_search(text))
        self.search_timer.start(self.config.get('behavior.search_delay_ms', 300))

    def perform_search(self, keyword: str):
        """执行搜索"""
        results = self.library.search_stickers(keyword)
        self.display_stickers(results)
        logging.info(f"搜索 '{keyword}' 找到 {len(results)} 个结果")

    def hide_window(self):
        """隐藏窗口"""
        self.hide()
        logging.info("窗口已隐藏")

    def show_window(self):
        """显示窗口"""
        self.show()
        self.activateWindow()
        self.raise_()
        logging.info("窗口已显示")

    def closeEvent(self, event):
        """关闭事件 - 隐藏而不是退出"""
        event.ignore()
        self.hide_window()


# ==================== 系统托盘 ====================
class SystemTrayManager(QObject):
    """系统托盘管理器"""

    def __init__(
        self,
        app: QApplication,
        window: StickerManagerWindow,
        config: ConfigManager,
        hotkey_listener: HotkeyListener,
    ):
        super().__init__()
        self.app = app
        self.window = window
        self.config = config
        self.hotkey_listener = hotkey_listener

        # 创建托盘图标
        self.tray_icon = QSystemTrayIcon(self.app)

        # 设置图标，增加错误处理
        try:
            if program_icon and os.path.exists(program_icon):
                icon = QIcon(program_icon)
                if not icon.isNull():
                    self.tray_icon.setIcon(icon)
                else:
                    # 使用默认图标
                    self.tray_icon.setIcon(
                        self.app.style().standardIcon(
                            self.app.style().StandardPixmap.SP_ComputerIcon
                        )
                    )
                    logging.warning("图标加载失败，使用默认图标")
            else:
                # 使用默认图标
                self.tray_icon.setIcon(
                    self.app.style().standardIcon(
                        self.app.style().StandardPixmap.SP_ComputerIcon
                    )
                )
                logging.warning("图标文件不存在，使用默认图标")
        except Exception as e:
            logging.error(f"设置托盘图标失败: {e}")
            # 使用默认图标
            self.tray_icon.setIcon(
                self.app.style().standardIcon(
                    self.app.style().StandardPixmap.SP_ComputerIcon
                )
            )

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = tray_menu.addAction("显示窗口")
        show_action.triggered.connect(self.window.show_window)

        reload_action = tray_menu.addAction("重新加载")
        reload_action.triggered.connect(self.reload_library)

        tray_menu.addSeparator()

        quit_action = tray_menu.addAction("退出")
        quit_action.triggered.connect(self.quit_app)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)

        self.tray_icon.show()
        logging.info("系统托盘已创建")

    def on_tray_activated(self, reason):
        """托盘图标激活事件"""
        if reason == QSystemTrayIcon.DoubleClick:
            if self.window.isHidden():
                self.window.show_window()
            else :
                self.window.hide()

    def reload_library(self):
        """重新加载表情库"""
        self.window.library.load_library()
        self.window.populate_categories()
        logging.info("表情库已重新加载")

    def quit_app(self):
        """退出应用"""
        logging.info("应用退出中...")

        # 停止热键监听
        self.hotkey_listener.stop()

        # 保存配置
        self.config.save_config()

        # 退出应用
        self.app.quit()


# ==================== 主程序 ====================
def main():
    """主函数"""
    # 设置日志
    setup_logging()
    logging.info("=" * 50)
    logging.info("表情包管理器启动")
    logging.info("=" * 50)

    # 加载配置
    config = ConfigManager()

    # 创建应用
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # 关闭窗口不退出应用
    app.setWindowIcon(QIcon(program_icon))

    # 创建热键监听器
    hotkey = config.get('hotkey', 'ctrl+shift+e')
    hotkey_listener = HotkeyListener(hotkey)

    # 创建主窗口
    window = StickerManagerWindow(config)

    # 连接热键信号
    hotkey_listener.hotkey_pressed.connect(lambda: window.show_window() if window.isHidden() else window.hide_window())

    # 启动热键监听
    hotkey_listener.start()

    # 创建系统托盘
    tray = SystemTrayManager(app, window, config, hotkey_listener)

    # 显示窗口
    # window.show_window()

    # 运行应用
    exit_code = app.exec()

    # 清理
    hotkey_listener.stop()

    logging.info("应用已退出")
    sys.exit(exit_code)


if __name__ == '__main__':
    main()
