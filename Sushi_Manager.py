
import os
import sys
import threading
import time
import re
import copy
import csv
import customtkinter as ctk
from PIL import Image, ImageTk, ImageDraw, ImageFont
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from collections import defaultdict
from dateutil.parser import parse
import traceback

# ======== 添加资源路径处理函数 ========
def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        # PyInstaller创建的临时文件夹
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

# ========== Configuration ==========
# 修改LOGO_PATH引用，使用resource_path函数
LOGO_PATH = resource_path("SELOGO22 - 01.png")
PASSWORD = "OPS123"
VERSION = "2.6.1"
DEVELOPER = "OPS - Voon Kee"

# Enhanced Theme Configuration
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ========== Enhanced Color Definitions with RGBA ==========
DARK_BG = "#0f172a"         # Original: (15, 23, 42, 0.8)
DARK_PANEL = "#1e293b"      # Original: (30, 41, 59, 0.8)
ACCENT_BLUE = "#3b82f6"     # Original: (59, 130, 246, 0.8)
BTN_HOVER = "#60a5fa"       # Original: (96, 165, 250, 0.8)
ACCENT_GREEN = "#10b981"    # Original: (16, 185, 129, 0.8)
ACCENT_RED = "#ef4444"      # Original: (239, 68, 68, 0.8)
ACCENT_PURPLE = "#8b5cf6"   # Original: (139, 92, 246, 0.8)
ENTRY_BG = "#334155"        # Original: (51, 65, 85, 0.8)
TEXT_COLOR = "#e2e8f0"
PANEL_BG = "#1e293b"        # Same as DARK_PANEL
TEXTBOX_BG = "#0f172a"      # Same as DARK_BG

# ========== Font Configuration (English in Italic) ==========
FONT_TITLE = ("Microsoft JhengHei", 24, "bold")        # 中文标题（不变）
FONT_BIGBTN = ("Microsoft JhengHei", 16, "bold")      # 中文按钮（不变）
FONT_MID = ("Microsoft JhengHei", 14)                 # 中文正文（不变）
FONT_SUB = ("Microsoft JhengHei", 12)                 # 中文辅助文字（不变）
FONT_ZH = ("Microsoft JhengHei", 12)                  # 中文专用字体（不变）
FONT_EN = ("Segoe UI", 11, "italic")                  # 英文改为斜体（添加 "italic"）
FONT_LOG = ("Consolas", 14)                           # 日志字体（不变）

# ========== Utility Functions ==========
def t(text):
    """Bilingual translation function"""
    translations = {
        "mapping_not_available": "分店供应商对应数据不可用\nOutlet-supplier mapping not available",
        "log_not_available": "日志数据不可用\nLog data not available",
        "info": "信息\nInformation",
        "processing": "處理中...\nProcessing...",
        "please_wait": "請稍候...\nPlease wait...",
        "error": "錯誤\nError",
        "login": "系統登錄\nSystem Login",
        "password": "輸入密碼...\nEnter password...",
        "login_btn": "登入\nLogin",
        "exit_confirm": "確定要退出應用程序嗎？\nAre you sure you want to exit the application?",
        "incorrect_pw": "密碼不正確，請重試\nIncorrect password, please try again",
        "main_title": "Sushi Express 自動化工具\nSushi Express Automation Tool",
        "select_function": "請選擇要執行的功能\nPlease select a function",
        "download_title": "Outlook 訂單下載\nOutlook Order Download",
        "download_desc": "下載本週的 Weekly Order 附件\nDownload weekly order attachments",
        "browse": "瀏覽...\nBrowse...",
        "start_download": "開始下載\nStart Download",
        "back_to_menu": "返回主菜單\nBack to Main Menu",
        "checklist_title": "Weekly Order 檢查表\nWeekly Order Checklist",
        "checklist_desc": "請選擇包含供應商訂單的資料夾\nSelect folder with supplier orders",
        "run_check": "執行檢查\nRun Check",
        "automation_title": "訂單自動彙整\nOrder Automation",
        "automation_desc": "請選擇三個必要的資料夾\nSelect required folders",
        "source_folder": "來源資料夾 (Weekly Orders)\nSource Folder (Weekly Orders)",
        "supplier_folder": "供應商資料夾 (Supplier)\nSupplier Folder",
        "outlet_folder": "分店資料夾 (Outlet)\nOutlet Folder",
        "start_automation": "開始彙整\nStart Automation",
        "processing_orders": "處理訂單\nProcessing Orders",
        "outlet_suppliers": "分店供應商對應\nOutlet-Supplier Mapping",
        "exit_system": "退出系統\nExit System",
        "select_account": "選擇 Outlook 帳號\nSelect Outlook Account",
        "enter_index": "請輸入序號：\nPlease enter index:",
        "download_summary": "下載摘要\nDownload Summary",
        "auto_download": "自動下載\nAuto Downloaded",
        "skipped": "跳過\nSkipped",
        "saved_to": "保存到\nSaved to",
        "check_results": "檢查結果\nCheck Results",
        "success": "成功\nSuccess",
        "warning": "警告\nWarning",
        "folder_warning": "請先選擇所有必要的資料夾\nPlease select all required folders",
        "close": "關閉\nClose",
        "order_processing": "訂單處理進度\nOrder Processing Progress",
        "outlet_supplier_mapping": "分店-供應商對應關係\nOutlet-Supplier Mapping",
        "select_folder": "選擇文件夾\nSelect Folder",
        "view_mapping": "查看分店供應商對應\nView Outlet-Supplier Mapping",
        "view_log": "查看完整日誌\nView Full Log",
        "supplier_files": "已處理的供應商文件\nProcessed Supplier Files",
        "outlet_files": "已處理的分店文件\nProcessed Outlet Files",
        "outlet_orders": "分店訂購情況\nOutlet Orders",
        "supplier_orders": "供應商訂購情況\nSupplier Orders",
    }
    return translations.get(text, text)

def get_contrast_color(bg_color):
    """Get contrasting text color for given background"""
    if isinstance(bg_color, tuple) and len(bg_color) >= 3:
        r, g, b = bg_color[:3]
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        return "#191F2B" if luminance > 170 else "#F3F6FA"
    else:
        # Fallback for HEX colors
        bg_color = bg_color.lstrip('#')
        r = int(bg_color[0:2], 16)
        g = int(bg_color[2:4], 16)
        b = int(bg_color[4:6], 16)
        luminance = 0.299 * r + 0.587 * g + 0.114 * b
        return "#191F2B" if luminance > 170 else "#F3F6FA"

def load_image(path, max_size=(400, 130)):
    """Safely load image with error handling"""
    try:
        if not os.path.exists(path):
            # Create a blank image if logo is missing
            img = Image.new('RGB', max_size, color=DARK_BG[:3])
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("arial.ttf", 24)
            draw.text((10,10), "Logo Missing", fill="white", font=font)
        else:
            img = Image.open(path)
        img.thumbnail(max_size, Image.LANCZOS)
        return ctk.CTkImage(img, size=img.size)
    except Exception as e:
        print(f"Error loading image: {e}")
        return None

# ========== Custom UI Components ==========
class GlowButton(ctk.CTkButton):
    """Button with glow effect and enhanced styling"""
    def __init__(self, master, text=None, glow_color=ACCENT_BLUE, **kwargs):
        super().__init__(master, text=text, **kwargs)
        self._glow_color = glow_color
        self._setup_style()
        self._bind_events()

    def _setup_style(self):
        self.configure(
            border_width=0,
            fg_color=self._glow_color,
            hover_color=self._adjust_color(self._glow_color, 20),
            text_color=get_contrast_color(self._glow_color),
            corner_radius=12,
            font=FONT_BIGBTN,
            height=50
        )

    def _bind_events(self):
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, event=None):
        self.configure(border_width=3, border_color=self._adjust_color(self._glow_color, 40))

    def _on_leave(self, event=None):
        self.configure(border_width=0)
    
    @staticmethod
    def _adjust_color(color, amount):
        """Adjust color brightness for RGBA tuples"""
        if isinstance(color, tuple) and len(color) >= 3:
            # For RGBA tuples
            r, g, b = color[:3]
            adjusted = tuple(min(255, max(0, x + amount)) for x in (r, g, b))
            if len(color) == 4:
                return adjusted + (color[3],)  # Keep alpha
            return adjusted
        else:
            # Fallback for HEX colors
            color = color.lstrip('#')
            rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            adjusted = tuple(min(255, max(0, x + amount)) for x in rgb)
            return f"#{adjusted[0]:02x}{adjusted[1]:02x}{adjusted[2]:02x}"

class ProgressPopup(ctk.CTkToplevel):
    """Popup window to show processing progress"""
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self.parent = parent
        self.configure(fg_color=DARK_BG)
        
        # Create log display
        self.log_text = ctk.CTkTextbox(
            self,
            wrap="word",
            font=FONT_LOG,
            fg_color=TEXTBOX_BG,
            text_color=TEXT_COLOR,
            corner_radius=10
        )
        self.log_text.pack(fill="both", expand=True, padx=20, pady=20)
        self.log_text.configure(state="disabled")
        
        # Close button
        close_btn = GlowButton(
            self,
            text=t("close"),
            command=self.destroy_popup,
            width=120,
            height=40
        )
        close_btn.pack(pady=10)
    
    def destroy_popup(self):
        """Safely destroy the popup"""
        self.destroy()
        self.parent.progress_popup = None
    
    def log(self, message):
        """Add message to log"""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

class MappingPopup(ctk.CTkToplevel):
    """Popup window to show outlet-supplier mapping"""
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self.parent = parent
        self.configure(fg_color=DARK_BG)
        
        # Create mapping display
        self.mapping_text = ctk.CTkTextbox(
            self,
            wrap="word",
            font=FONT_LOG,
            fg_color=TEXTBOX_BG,
            text_color=TEXT_COLOR,
            corner_radius=10
        )
        self.mapping_text.pack(fill="both", expand=True, padx=20, pady=20)
        self.mapping_text.configure(state="disabled")
        
        # Close button
        close_btn = GlowButton(
            self,
            text=t("close"),
            command=self.destroy_popup,
            width=120,
            height=40
        )
        close_btn.pack(pady=10)
    
    def destroy_popup(self):
        """Safely destroy the popup"""
        self.destroy()
        self.parent.mapping_popup = None
    
    def update_mapping(self, mapping):
        """Update mapping display"""
        self.mapping_text.configure(state="normal")
        self.mapping_text.delete("1.0", "end")
        self.mapping_text.insert("1.0", mapping)
        self.mapping_text.configure(state="disabled")

class ScrollableMessageBox(ctk.CTkToplevel):
    """Scrollable message box for displaying long text"""
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self.configure(fg_color=DARK_BG)
        
        # Create text box
        self.text_box = ctk.CTkTextbox(
            self,
            wrap="word",
            font=FONT_LOG,
            fg_color=TEXTBOX_BG,
            text_color=TEXT_COLOR,
            corner_radius=10
        )
        self.text_box.pack(fill="both", expand=True, padx=20, pady=20)
        self.text_box.insert("1.0", message)
        self.text_box.configure(state="disabled")
        
        # Close button
        close_btn = GlowButton(
            self,
            text=t("close"),
            command=self.destroy,
            width=120,
            height=40
        )
        close_btn.pack(pady=10)

# ========== 送货日期验证工具 ==========
class DeliveryDateValidator:
    """送货日期验证工具（支持分店特定规则）"""
    
    DAYS_MAPPING = {
        'mon': 0, 'monday': 0, '星期一': 0,
        'tue': 1, 'tuesday': 1, '星期二': 1,
        'wed': 2, 'wednesday': 2, '星期三': 2,
        'thu': 3, 'thursday': 3, '星期四': 3,
        'fri': 4, 'friday': 4, '星期五': 4,
        'sat': 5, 'saturday': 5, '星期六': 5,
        'sun': 6, 'sunday': 6, '星期日': 6
    }

    def __init__(self, config_file=None):
        self.schedule = defaultdict(dict)
        if config_file:
            self.load_config(config_file)
    
    def load_config(self, config_file):
        """加载供应商-分店送货日期配置"""
        try:
            with open(config_file, mode='r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    supplier = row['supplier'].strip()
                    outlet = row['outlet_code'].strip().upper()
                    days = self.parse_delivery_days(row['delivery_days'])
                    
                    if outlet == "ALL":
                        # 适用于该供应商所有分店
                        self.schedule[supplier]['*'] = days
                    else:
                        self.schedule[supplier][outlet] = days
        except Exception as e:
            raise Exception(f"加载送货配置失败: {str(e)}")

    def parse_delivery_days(self, days_str):
        """解析送货日期字符串"""
        days = set()
        for day in days_str.split(','):
            day = day.strip().lower()
            if day in self.DAYS_MAPPING:
                days.add(self.DAYS_MAPPING[day])
        return days

    def get_delivery_days(self, supplier, outlet_code):
        """获取特定供应商-分店的送货日"""
        # 先检查是否有分店特定规则
        outlet_specific = self.schedule.get(supplier, {}).get(outlet_code)
        if outlet_specific is not None:
            return outlet_specific
        
        # 检查是否有全局规则
        return self.schedule.get(supplier, {}).get('*', set())

    def validate_order(self, supplier, outlet_code, order_date, log_callback=None):
        """验证订单日期"""
        delivery_days = self.get_delivery_days(supplier, outlet_code)
        
        if not delivery_days:  # 无限制则自动通过
            return True
        
        try:
            # 解析订单日期
            if isinstance(order_date, (int, float)):  # Excel日期值
                order_date = datetime(1899, 12, 30) + timedelta(days=order_date)
            else:
                order_date = parse(str(order_date), fuzzy=True)
            
            order_weekday = order_date.weekday()
            
            if order_weekday not in delivery_days:
                day_name = ['Monday', 'Tuesday', 'Wednesday', 'Thursday',
                           'Friday', 'Saturday', 'Sunday'][order_weekday]
                if log_callback:
                    log_callback(
                        f"❌ 送货日期错误: {outlet_code} 向 {supplier} 下单\n"
                        f"  订单日期: {order_date.strftime('%Y-%m-%d')} ({day_name})\n"
                        f"  允许送货日: {self.format_days(delivery_days)}"
                    )
                return False
            return True
        except Exception as e:
            if log_callback:
                log_callback(f"⚠️ 日期解析失败: {supplier}-{outlet_code} ({order_date}): {str(e)}")
            return False

    def format_days(self, day_numbers):
        """将数字转换为星期名称"""
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 
               'Friday', 'Saturday', 'Sunday']
        return ", ".join(days[d] for d in sorted(day_numbers))

# ========== Order Automation Core ==========
class OrderAutomation:
    """订单自动化工具（已集成地址和日期检查）"""
    
    OUTLET_LIST = [
        # Sushi Express
        ("Sushi Express Century Square", "CSQ"),
        ("Sushi Express Clementi Mall", "TCM"),
        ("Sushi Express Funan", "FN"),
        ("Sushi Express Heartbeat @ Bedok", "HBB"),
        ("Sushi Express Heartland Mall", "HLM"),
        ("Sushi Express Hillion Mall", "HM"),
        ("Sushi Express Hougang Mall", "HGM"),
        ("Sushi Express IMM", "IMM"),
        ("Sushi Express Jurong Point", "JP"),
        ("Sushi Express NEX Serangoon", "NEX"),
        ("Sushi Express North Point", "NPC"),
        ("Sushi Express Parkway Parade", "PP"),
        ("Sushi Express Paya Lebar Quarter", "PLQ"),
        ("Sushi Express Sengkang Grand", "SKG"),
        ("Sushi Express Seletar Mall", "SM"),
        ("Sushi Express Sun Plaza", "SP"),
        ("Sushi Express Waterway Point", "WWP"),
        ("Sushi Express Westgate", "WG"),
        ("Sushi Express White Sands", "WS"),
        ("Sushi Express Central Kitchen", "CK"),
        ("Sushi Express West Mall", "WM"),
        # TakeOut
        ("Sushi TakeOut CityVibe (GTM)", "GTM"),
        ("Sushi TakeOut Tampines SMRT", "TSMRT"),
        ("Sushi TakeOut Woodlands", "WL"),
        ("Sushi TakeOut Toa Payoh", "TPY"),
        # Sushi GOGO
        ("Sushi GOGO Junction 8", "J8"),
        ("Sushi GOGO Hougang MRT", "HGTO"),  # Fixed HGTO mapping
        ("Sushi GOGO Oasis Terraces", "OASIS"),
        ("Sushi GOGO Sengkang MRT", "SKMRT"),
        ("Sushi GOGO Yew Tee Square", "YTS"),
        ("Sushi GOGO The Poiz Centre", "TPC"),
        ("Sushi GOGO Ang Mo Kio Hub", "AMK"),
        ("Sushi GOGO Canberra Plaza", "CBP"),
        ("Sushi GOGO Bukit Gombak", "BGM"),
        ("Sushi GOGO Pasir Ris Mall", "PRM"),
        ("Sushi GOGO North Point City", "NPG"),
        # SushiPlus
        ("SushiPlus Bugis", "Bugis"),
        ("SushiPlus 313 Somerset", "313"),
        ("SushiPlus Tampines One", "T1"),
    ]

    @staticmethod
    def get_short_code(f5val):
        """Enhanced outlet code matching with HGTO fix"""
        val = (str(f5val) or "").strip().lower()
        
        # 1. Exact match for full name
        for full, code in OrderAutomation.OUTLET_LIST:
            if full.lower() == val:
                return code
        
        # 2. Partial match (contains)
        for full, code in OrderAutomation.OUTLET_LIST:
            if full.lower() in val:
                return code
        
        # 3. Keyword match (last 2 keywords)
        for full, code in OrderAutomation.OUTLET_LIST:
            key_words = [x.lower() for x in full.split() if len(x) > 1]
            if key_words and all(kw in val for kw in key_words[-2:]):
                return code
        
        # 4. Code match
        for full, code in OrderAutomation.OUTLET_LIST:
            if code.lower() in val:
                return code
        
        # 5. Special handling for common issues
        special_cases = {
            "hougang mrt": "HGTO",
            "hgto": "HGTO",
            "gogo hougang": "HGTO",
            "sushi gogo hougang": "HGTO",
            "west mall": "WM",
            "westmall": "WM",
            "wsm": "WM"
        }
        
        for pattern, code in special_cases.items():
            if pattern in val:
                return code
        
        # 6. Generate simplified code
        code = "".join([x for x in val if x.isalnum()])
        return code[:4].upper() if code else "UNKNOWN"

    @staticmethod
    def is_valid_date(cell_value, next_week_start, next_week_end):
        """Check if cell value is a valid date within the target week"""
        try:
            # Handle Excel serial dates
            if isinstance(cell_value, (int, float)):
                base_date = datetime(1899, 12, 30)
                parsed = base_date + timedelta(days=cell_value)
            else:
                parsed = parse(str(cell_value), fuzzy=True, dayfirst=False)
            
            return next_week_start.date() <= parsed.date() <= next_week_end.date()
        except:
            return False

    @classmethod
    def find_delivery_date_row(cls, ws, next_week_start, next_week_end, max_rows=150):
        """Find row with delivery dates in columns F-K"""
        valid_col_range = range(5, 12)  # Columns F to K
        invalid_labels = ["total", "total:", "sub-total", "sub-total:", "no. of cartons", "no. of cartons:"]
        
        found_blocks = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows)):
            cols = []
            for j, cell in enumerate(row):
                if j not in valid_col_range:
                    continue
                if cls.is_valid_date(cell.value, next_week_start, next_week_end):
                    cols.append(j)
            if cols:
                found_blocks.append((i + 1, cols))
        
        for header_row, cols in found_blocks:
            for row_idx in range(header_row + 1, header_row + 100):
                label_cell = ws.cell(row=row_idx, column=5)  # Column E
                label = str(label_cell.value).lower().strip() if label_cell.value else ""
                
                if any(invalid in label for invalid in invalid_labels):
                    continue
                
                for col_idx in cols:
                    qty_cell = ws.cell(row=row_idx, column=col_idx)
                    if isinstance(qty_cell.value, (int, float)) and qty_cell.value > 0:
                        return header_row, cols
        
        return None, []

    @classmethod
    def run_automation(cls, source_folder, supplier_folder, outlet_folder, 
                      outlet_config=None, delivery_config=None,
                      log_callback=None, mapping_callback=None):
        """Complete order automation with exact formatting and HGTO fix"""
        now = datetime.now()
        next_week_start = (now + timedelta(days=7 - now.weekday())).replace(hour=0, minute=0, second=0)
        next_week_end = next_week_start + timedelta(days=6)
        log_file = os.path.join(source_folder, f"order_log_{now.strftime('%Y%m%d_%H%M%S')}.txt")
        log_lines = [f"Scan Start Time: {now}\n", f"Target Week: {next_week_start.date()} to {next_week_end.date()}\n"]
        
        # 初始化日期验证器
        date_validator = DeliveryDateValidator(delivery_config) if delivery_config else None
        
        # 准备分店配置
        outlet_mapping = {}
        if outlet_config:
            try:
                outlet_mapping = {o['short_name']: o for o in outlet_config}
                log_callback(f"✅ 已加载分店配置: {len(outlet_mapping)} 个分店")
            except Exception as e:
                log_callback(f"⚠️ 分店配置加载失败: {str(e)}")
        
        # Store outlet-supplier relationships
        supplier_to_outlets = defaultdict(list)
        outlet_to_suppliers = defaultdict(list)
        
        # Ensure directories exist
        os.makedirs(supplier_folder, exist_ok=True)
        os.makedirs(outlet_folder, exist_ok=True)
        
        # Log callback function
        def log(message, include_timestamp=True):
            timestamp = datetime.now().strftime("%H:%M:%S") if include_timestamp else ""
            log_line = f"[{timestamp}] {message}" if include_timestamp else message
            log_lines.append(log_line + "\n")
            if log_callback:
                log_callback(log_line + "\n")
        
        # Mapping callback function
        def update_mapping():
            if mapping_callback:
                mapping_text = "分店-供应商对应关系:\n"
                for outlet, suppliers in outlet_to_suppliers.items():
                    mapping_text += f"{outlet}: {', '.join(suppliers)}\n"
                mapping_callback(mapping_text)
        
        # Process source folder files
        files = [f for f in os.listdir(source_folder) 
                 if f.endswith((".xlsx", ".xls")) and not f.startswith("~$")]
        total_files = len(files)
        
        if total_files == 0:
            log("No Excel files found in source folder. Exiting.")
            return False, "在源文件夹中未找到Excel文件"
        
        log(f"找到 {total_files} 个Excel文件")
        log(f"目标周期: {next_week_start.strftime('%Y-%m-%d')} 到 {next_week_end.strftime('%Y-%m-%d')}")
        
        for idx, file in enumerate(files):
            full_path = os.path.join(source_folder, file)
            try:
                log(f"\n处理文件 {idx+1}/{total_files}: {file}")
                wb = load_workbook(full_path, data_only=True)
                
                # 记录文件中的所有工作表
                log(f"工作表: {', '.join(wb.sheetnames)}")
                
                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    if ws.sheet_state != "visible":
                        log(f"  ⏩ 跳过隐藏工作表: {sheetname}")
                        continue
                    
                    # 获取分店代码
                    f5_value = ws['F5'].value if 'F5' in ws else ws.cell(row=5, column=6).value
                    outlet_short = cls.get_short_code(f5_value)
                    log(f"  Sheet: {sheetname}, F5 Value: '{f5_value}', Short Code: {outlet_short}")
                    
                    # ====== 新增: 地址一致性检查 ======
                    if outlet_mapping:
                        outlet_info = outlet_mapping.get(outlet_short)
                        if outlet_info:
                            # 获取工作表地址 (F6单元格)
                            sheet_address = ws['F6'].value if 'F6' in ws else None
                            if sheet_address:
                                # 标准化地址比较
                                def normalize(addr):
                                    return re.sub(r'[^a-zA-Z0-9]', '', str(addr).lower())
                                
                                config_addr = normalize(outlet_info.get('address', ''))
                                sheet_addr = normalize(sheet_address)
                                
                                if config_addr != sheet_addr:
                                    log(f"  ❌ 地址不匹配: {outlet_short}")
                                    log(f"    配置地址: {outlet_info.get('address', '')}")
                                    log(f"    工作表地址: {sheet_address}")
                                    # 标记问题但不停止处理
                                    outlet_short = f"{outlet_short} (地址错误)"
                    
                    # ====== 新增: 送货日期检查 ======
                    if date_validator:
                        # 获取送货日期 (F8单元格)
                        delivery_date = ws['F8'].value if 'F8' in ws else None
                        if delivery_date:
                            is_valid = date_validator.validate_order(
                                sheetname, outlet_short, delivery_date, log
                            )
                            if not is_valid:
                                log(f"  ⏩ 跳过 {outlet_short} 的订单（日期不符合要求）")
                                continue
                        else:
                            log(f"  ⚠️ {outlet_short}: 未找到送货日期(F8)")
                    
                    # 查找送货日期行
                    header_row, target_cols = cls.find_delivery_date_row(ws, next_week_start, next_week_end)
                    
                    has_order = False
                    if target_cols:
                        log(f"    Found delivery dates at row {header_row}, columns {[get_column_letter(c) for c in target_cols]}")
                        
                        # 检查订单数量
                        for row_idx in range(header_row + 1, header_row + 100):
                            label_cell = ws.cell(row=row_idx, column=5)  # Column E
                            label = str(label_cell.value).lower().strip() if label_cell.value else ""
                            
                            if any(invalid in label for invalid in ["total", "sub-total", "no. of cartons"]):
                                continue
                            
                            for col_idx in target_cols:
                                qty_cell = ws.cell(row=row_idx, column=col_idx)
                                if isinstance(qty_cell.value, (int, float)) and qty_cell.value > 0:
                                    has_order = True
                                    break
                            
                            if has_order:
                                break
                    else:
                        log("    No valid delivery dates found in columns F-K")
                    
                    if has_order:
                        log(f"    ✅ {outlet_short} has orders for {sheetname}")
                        supplier_to_outlets[sheetname].append((outlet_short, full_path, sheetname))
                        outlet_to_suppliers[outlet_short].append(sheetname)
                    else:
                        log(f"    ⏩ No orders found for {outlet_short} in {sheetname}")
                
            except Exception as e:
                error_msg = f"❌ Error processing {file}: {str(e)}\n{traceback.format_exc()}"
                log(error_msg)
        
        # 记录分店-供应商关系
        log("\nOutlet-Supplier Mapping:")
        mapping_text = ""
        for outlet, suppliers in outlet_to_suppliers.items():
            log(f"  {outlet}: {', '.join(suppliers)}")
            mapping_text += f"{outlet}: {', '.join(suppliers)}\n"
        
        # 更新映射回调
        if mapping_callback:
            mapping_callback(mapping_text)

        log("\nSupplier-Outlet Mapping:")
        for supplier, outlets in supplier_to_outlets.items():
            outlet_list = [o[0] for o in outlets]
            log(f"  {supplier}: {', '.join(outlet_list)}")

        # SUPPLIER 合併 - 创建供应商文件
        log("\nCreating supplier files...")
        supplier_files = []
        for sheetname, outlet_file_pairs in supplier_to_outlets.items():
            supplier_path = os.path.join(supplier_folder, f"{sheetname}_Week_{next_week_start.strftime('%V')}.xlsx")
            new_wb = Workbook()
            new_wb.remove(new_wb.active)  # 删除默认sheet
            
            log(f"Creating supplier file for {sheetname} at {supplier_path}")
            
            for outlet, src_file, original_sheet in outlet_file_pairs:
                try:
                    log(f"  Adding outlet: {outlet} from {src_file}")
                    src_wb = load_workbook(src_file, data_only=False)
                    src_ws = src_wb[original_sheet]
                    
                    # 跳过隐藏的工作表
                    if src_ws.sheet_state != "visible":
                        log(f"    ⏩ Skipping hidden sheet: {original_sheet}")
                        continue
                    
                    # 创建新工作表
                    target_ws = new_wb.create_sheet(title=outlet)
                    
                    # 复制所有行和列
                    for row in src_ws.iter_rows():
                        for cell in row:
                            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                            
                            # 复制样式
                            if cell.has_style:
                                new_cell.font = copy.copy(cell.font)
                                new_cell.border = copy.copy(cell.border)
                                new_cell.fill = copy.copy(cell.fill)
                                new_cell.number_format = cell.number_format
                                new_cell.protection = copy.copy(cell.protection)
                                new_cell.alignment = copy.copy(cell.alignment)
                    
                    # 复制列宽
                    for col_idx in range(1, src_ws.max_column + 1):
                        col_letter = get_column_letter(col_idx)
                        target_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                    
                    # 复制行高
                    for row_idx in range(1, src_ws.max_row + 1):
                        target_ws.row_dimensions[row_idx].height = src_ws.row_dimensions[row_idx].height
                    
                    # 复制合并单元格
                    for merged_range in src_ws.merged_cells.ranges:
                        target_ws.merge_cells(str(merged_range))
                        
                except Exception as e:
                    error_msg = f"    ❌ Failed to copy {outlet} in {sheetname}: {str(e)}\n{traceback.format_exc()}"
                    log(error_msg)
            
            try:
                new_wb.save(supplier_path)
                log(f"  ✅ Saved supplier file: {supplier_path}")
                supplier_files.append(os.path.basename(supplier_path))
            except Exception as e:
                error_msg = f"    ❌ Failed to save {sheetname} file: {str(e)}"
                log(error_msg)

        # OUTLET 合併 - 创建分店文件
        log("\nCreating outlet files...")
        outlet_files = []
        outlet_to_sheets = defaultdict(list)
        for sheetname, pairs in supplier_to_outlets.items():
            for outlet, path, original_sheet in pairs:
                outlet_to_sheets[outlet].append((sheetname, path, original_sheet))

        for outlet, supplier_sheets in outlet_to_sheets.items():
            out_path = os.path.join(outlet_folder, f"{outlet}_Week_{next_week_start.strftime('%V')}.xlsx")
            out_wb = Workbook()
            out_wb.remove(out_wb.active)  # 删除默认sheet
            
            log(f"Creating outlet file for {outlet} at {out_path}")
            
            for supplier, src_file, original_sheet in supplier_sheets:
                try:
                    log(f"  Adding supplier: {supplier} from {src_file}")
                    src_wb = load_workbook(src_file, data_only=False)
                    src_ws = src_wb[original_sheet]
                    
                    # 创建新工作表（限制工作表名称为31字符）
                    sheet_title = supplier[:31]
                    target_ws = out_wb.create_sheet(title=sheet_title)
                    
                    # 复制所有行和列
                    for row in src_ws.iter_rows():
                        for cell in row:
                            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                            
                            # 复制样式
                            if cell.has_style:
                                new_cell.font = copy.copy(cell.font)
                                new_cell.border = copy.copy(cell.border)
                                new_cell.fill = copy.copy(cell.fill)
                                new_cell.number_format = cell.number_format
                                new_cell.protection = copy.copy(cell.protection)
                                new_cell.alignment = copy.copy(cell.alignment)
                    
                    # 复制列宽
                    for col_idx in range(1, src_ws.max_column + 1):
                        col_letter = get_column_letter(col_idx)
                        target_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                    
                    # 复制行高
                    for row_idx in range(1, src_ws.max_row + 1):
                        target_ws.row_dimensions[row_idx].height = src_ws.row_dimensions[row_idx].height
                    
                    # 复制合并单元格
                    for merged_range in src_ws.merged_cells.ranges:
                        target_ws.merge_cells(str(merged_range))
                        
                except Exception as e:
                    error_msg = f"    ❌ Failed to copy {supplier} to {outlet}: {str(e)}\n{traceback.format_exc()}"
                    log(error_msg)
            
            try:
                out_wb.save(out_path)
                log(f"  ✅ Saved outlet file: {out_path}")
                outlet_files.append(os.path.basename(out_path))
            except Exception as e:
                error_msg = f"    ❌ Failed to save outlet file for {outlet}: {str(e)}"
                log(error_msg)

        # 添加分店文件列表到日志
        result_text = "\n订单整合结果:\n"
        result_text += f"已处理供应商文件: {len(supplier_files)}\n"
        result_text += f"已处理分店文件: {len(outlet_files)}\n\n"
        
        result_text += "供应商文件列表:\n" + "\n".join([f"- {file}" for file in supplier_files]) + "\n\n"
        result_text += "分店文件列表:\n" + "\n".join([f"- {file}" for file in outlet_files])
        
        log(result_text)
        
        # 保存日志文件
        try:
            with open(log_file, "w", encoding="utf-8") as logfile:
                logfile.writelines(log_lines)
            log(f"\nLog saved at: {log_file}")
            return True, f"订单整合完成！\n\n日志文件保存至:\n{log_file}\n\n{result_text}"
        except Exception as e:
            log(f"❌ Failed to write log file: {e}")
            return False, f"订单整合完成但日志保存失败:\n{str(e)}"

# ========== Order Checker Core ==========
class OrderChecker:
    """Order checklist tool with enhanced output"""
    @staticmethod
    def get_outlet_shortname(f5_value):
        """Get outlet short name from Excel F5 value"""
        if not f5_value or not isinstance(f5_value, str):
            return "[EMPTY]"

        f5_value = f5_value.lower().strip()
        mapping = [
            (r"gogo.*north point", "NPG"),
            (r"express.*north point", "NPC"),
            (r"north point city", "NPG"),
            (r"jurong point", "JP"),
            (r"junction 8", "J8"),
            (r"ang mo kio", "AMK"),
            (r"bugis", "Bugis"),
            (r"funan", "FN"),
            (r"yew tee", "YTS"),
            (r"cityvibe", "GTM"),
            (r"canberra", "CBP"),
            (r"hougang mrt", "HGTO"),
            (r"hougang", "HGM"),
            (r"poiz", "TPC"),
            (r"sun", "SP"),
            (r"waterway", "WWP"),
            (r"west mall", "WM"),
            (r"westmall", "WM"),
            (r"wsm", "WM"),
            (r"pasir ris", "PRM"),
            (r"bukit", "BGM"),
            (r"plq|paya lebar", "PLQ"),
            (r"white sands", "WS"),
            (r"seletar", "SM"),
            (r"clementi", "TCM"),
            (r"tampines smrt", "TSMRT"),
            (r"woodlands", "WL"),
            (r"toa payoh", "TPY"),
            (r"sengkang mrt", "SKMRT"),
            (r"sengkang grand", "SKG"),
            (r"hillion", "HM"),
            (r"heartbeat", "HBB"),
            (r"oasis", "OAS"),
            (r"westgate", "WG"),
            (r"imm", "IMM"),
            (r"nex", "NEX"),
            (r"century square", "CSQ"),
            (r"parkway", "PP"),
            (r"heartland", "HLM"),
            (r"313", "313"),
            (r"tampines one", "T1"),
        ]
        
        for pattern, short in mapping:
            if re.search(pattern, f5_value):
                return short
        return f"[UNKNOWN] {f5_value}"

    @classmethod
    def run_checklist(cls, folder, delivery_config=None, log_callback=None):
        """Run supplier checklist with enhanced output"""
        # 初始化日期验证器
        date_validator = DeliveryDateValidator(delivery_config) if delivery_config else None
        
        supplier_keywords = {
            "Chang Cheng": ["chang cheng"],
            "Super Q": ["super q"],
            "Super Q Chawan": ["super q chaw", "chawan", "super q chawan"],
            "Oriental": ["oriental"],
            "Aries": ["aries"],
            "Evershine": ["evershine"],
            "Perfect Choice": ["perfect choice"]
        }
        
        must_have_outlets = {
            "Chang Cheng": ["GTM", "HGTO", "TPC", "YTS", "TSMRT", "CSQ", "HBB", "HLM", "PLQ", "WS", "SM",
                          "HGM", "WWP", "WG", "NEX", "SP", "CK", "BGM", "CBP", "HM", "Bugis", "PP", "FN", "WL", "PRM"],
            "Super Q": ["AMK", "OAS", "J8", "SKMRT", "SKG", "TCM", "NPC", "IMM", "JP", "TPY", "NPG", "WM"],
            "Super Q Chawan": ["AMK", "WG", "NPC", "SP", "SKG", "WWP", "HGM", "WL", "HGTO", "SKMRT", "TPC", "BGM",
                              "NPG", "TCM", "IMM", "J8", "JP", "OAS", "TSMRT", "TPY", "WM"],
            "Oriental": ["313", "Bugis", "HGTO", "TPC", "WL", "HLM", "PLQ", "HGM", "SKG", "T1"],
            "Aries": ["AMK", "J8", "TPY", "CSQ", "HBB", "PP", "WS", "SM", "WWP", "FN", "IMM", "JP",
                     "WG", "NEX", "OAS", "CK", "WM"],
            "Evershine": ["YTS", "GTM", "TCM", "SKMRT", "TSMRT", "HM", "NPC", "SP", "NPG"],
            "Perfect Choice": ["GTM", "CBP", "FN", "CSQ", "WS", "HBB", "PLQ", "PP", "HM", "HLM", "NEX", "SM", "YTS"]
        }
        
        try:
            files = [f for f in os.listdir(folder) if f.endswith(".xlsx") and not f.startswith("~$")]
            normalized_files = {cls._normalize(f): f for f in files}
            output = []
            
            for supplier, keywords in supplier_keywords.items():
                match = cls._find_supplier_file(normalized_files, keywords)
                if not match:
                    output.append(f"\n❌ {supplier} - Supplier file not found.")
                    continue
                
                result = cls._process_supplier_file(folder, match, must_have_outlets[supplier], date_validator, log_callback)
                output.extend(result)
            
            return "\n".join(output)
        except Exception as e:
            return f"❌ Error running checklist: {str(e)}\n{traceback.format_exc()}"

    @staticmethod
    def _normalize(text):
        """Normalize text for matching"""
        return re.sub(r'[\s\(\)]', '', text.lower())

    @classmethod
    def _find_supplier_file(cls, normalized_files, keywords):
        """Find supplier file"""
        for k in keywords:
            nk = cls._normalize(k)
            for nf, of in normalized_files.items():
                if nk in nf:
                    return of
        return None

    @classmethod
    def _process_supplier_file(cls, folder, filename, required_outlets, date_validator=None, log_callback=None):
        """Process supplier file with date validation"""
        output = []
        found = set()
        unidentified = []
        
        try:
            wb = load_workbook(os.path.join(folder, filename), data_only=True)
            for s in wb.sheetnames:
                try:
                    f5 = wb[s]["F5"].value
                    code = cls.get_outlet_shortname(f5)
                    if "[UNKNOWN]" in code or "[EMPTY]" in code:
                        unidentified.append((s, f5))
                    else:
                        found.add(code)
                    
                    # ====== 新增: 送货日期检查 ======
                    if date_validator:
                        # 获取送货日期 (F8单元格)
                        delivery_date = wb[s]['F8'].value if 'F8' in wb[s] else None
                        if delivery_date:
                            is_valid = date_validator.validate_order(
                                os.path.basename(filename).split('_')[0],
                                code,
                                delivery_date,
                                log_callback
                            )
                            if not is_valid:
                                output.append(f"❌ 日期错误: {code} 的订单日期不符合要求")
                except:
                    unidentified.append((s, "[F5 error]"))
            
            required = set(required_outlets)
            output.append(f"\n=== {filename} ===")
            output.append(f"📊 Required: {len(required)}, Found: {len(found)}, Missing: {len(required - found)}")
            
            for o in sorted(required & found):
                output.append(f"✔️ {o}")
            for o in sorted(required - found):
                output.append(f"❌ {o}")
            for s, v in unidentified:
                output.append(f"⚠️ {s} => {v}")
                
        except Exception as e:
            output.append(f"❌ Error processing {filename}: {str(e)}")
        
        return output

# ========== Outlook Downloader Core ==========
class OutlookDownloader:
    """Outlook order downloader with enhanced UI"""
    OUTLET_MAP = {
        "century": "CSQ", "clementi": "TCM", "funan": "FN", "heartbeat": "HBB",
        "heartland": "HLM", "hillion": "HM", "hougang": "HGM", "imm": "IMM",
        "jurong": "JP", "nex": "NEX", "north point city": "NPG", "north point": "NPC",
        "parkway": "PP", "paya lebar": "PLQ", "sengkang grand": "SKG", "seletar": "SM",
        "sun plaza": "SP", "waterway": "WWP", "westgate": "WG", "white sands": "WS",
        "cityvibe": "GTM", "tampines smrt": "TSMRT", "woodlands": "WL", "toa payoh": "TPY",
        "junction 8": "J8TO", "hougang mrt": "HGTO", "oasis": "OASIS", "sengkang mrt": "SKMRT",
        "yew tee": "YTS", "poiz": "TPC", "central kitchen": "CK GOGO", "bugis": "Bugis",
        "313": "Sushi+ 313", "tampines one": "T1", "ang mo kio": "AMK", "canberra": "CBP",
        "gombak": "BKG", "pasir ris": "PRM", "lavendar": "VHL", "west mall": "WSM"
    }
    
    @classmethod
    def read_outlet_config(cls, config_file):
        """读取分店配置文件"""
        outlets = []
        try:
            with open(config_file, mode='r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    outlets.append({
                        'email': row['email'].strip().lower(),
                        'short_name': row['short_name'].strip(),
                        'full_name': row.get('full_name', '').strip(),
                        'address': row.get('address', '').strip(),
                        'delivery_day': row.get('delivery_day', '').strip(),
                    })
            return outlets
        except Exception as e:
            raise Exception(f"读取分店配置文件失败: {str(e)}")

    @classmethod
    def download_weekly_orders(cls, destination_folder, config_file=None, account_idx=None, callback=None):
        """Download weekly orders with enhanced error handling"""
        try:
            import win32com.client
            from win32com.client import Dispatch
        except ImportError:
            error_msg = "需要安裝 win32com 庫來使用 Outlook 功能\n請運行: pip install pywin32"
            if callback:
                callback(error_msg)
            else:
                messagebox.showerror(t("error"), error_msg)
            return

        # 读取分店配置
        email_to_outlet = {}
        if config_file:
            try:
                outlets = cls.read_outlet_config(config_file)
                email_to_outlet = {o['email']: o for o in outlets}
                if callback:
                    callback(f"✅ 已加载分店配置: {len(outlets)} 个分店")
            except Exception as e:
                if callback:
                    callback(f"❌ 分店配置读取失败: {str(e)}")
                return

        # 创建下载目录
        week_no = datetime.now().isocalendar()[1]
        save_path = os.path.join(destination_folder, f"Week_{week_no}")
        os.makedirs(save_path, exist_ok=True)

        # 设置时间范围（上周六到这周日）
        today = datetime.now()
        monday_this_week = today - timedelta(days=today.weekday())
        start_of_range = monday_this_week - timedelta(days=2)  # 上周六
        end_of_range = monday_this_week + timedelta(days=6, hours=23, minutes=59)

        # 初始化Outlook
        try:
            outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        except Exception as e:
            error_msg = f"無法啟動 Outlook: {str(e)}"
            if callback:
                callback(error_msg)
            return

        # 如果没有提供账号索引，让用户选择
        if account_idx is None:
            accounts = [outlook.Folders.Item(i + 1) for i in range(outlook.Folders.Count)]
            account_names = [acct.Name for acct in accounts]
            
            root = ctk.CTk()
            root.withdraw()
            account_idx = simpledialog.askinteger(
                t("select_account"),
                "\n".join([f"[{i}] {name}" for i, name in enumerate(account_names)]) + "\n" + t("enter_index"),
                minvalue=0, maxvalue=len(account_names)-1,
                parent=root
            )
            root.destroy()
            
            if account_idx is None:
                return
    
        try:
            account_folder = outlook.Folders.Item(account_idx + 1)
            messages = cls._collect_messages(account_folder, start_of_range, end_of_range)
            latest_messages = cls._filter_latest_messages(messages, email_to_outlet)
            result = cls._download_attachments(latest_messages, save_path, email_to_outlet, week_no)
            
            summary = (
                f"=== {t('download_summary')} ===\n"
                f"📅 日期范围: {start_of_range.strftime('%Y-%m-%d')} 至 {end_of_range.strftime('%Y-%m-%d')}\n"
                f"✅ {t('auto_download')}: {result['downloaded']}\n"
                f"⏩ {t('skipped')}: {result['skipped']}\n"
                f"📁 {t('saved_to')}: {save_path}\n\n"
                f"匹配的分店: {len(result['matched_outlets'])}\n"
                f"未匹配的邮件: {len(result['unmatched_emails'])}"
            )
            
            if result['unmatched_emails']:
                summary += f"\n\n未匹配的邮箱:\n" + "\n".join(result['unmatched_emails'])
            
            if callback:
                callback(summary)
        except Exception as e:
            error_msg = f"下載過程中出錯: {str(e)}"
            if callback:
                callback(error_msg)

    @classmethod
    def _collect_messages(cls, folder, start_date, end_date):
        """收集指定日期范围内的邮件"""
        messages = []
        try:
            # 创建Outlook日期格式的过滤字符串
            filter_str = (
                f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %I:%M %p')}' AND "
                f"[ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %I:%M %p')}'"
            )
            
            items = folder.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)  # 按接收时间降序排序
            messages.extend([msg for msg in items if msg.Class == 43])  # 43是邮件类
        except Exception as e:
            print(f"Error accessing folder {folder.Name}: {e}")
        
        # 递归检查子文件夹
        for sub in folder.Folders:
            messages.extend(cls._collect_messages(sub, start_date, end_date))
        
        return messages

    @classmethod
    def _filter_latest_messages(cls, messages, email_to_outlet=None):
        """过滤邮件，确保每个分店只保留最新的邮件"""
        latest_messages = {}
        
        for msg in messages:
            try:
                sender_email = msg.SenderEmailAddress.lower()
                
                # 尝试匹配分店
                outlet_key = sender_email
                if email_to_outlet:
                    outlet_info = email_to_outlet.get(sender_email)
                    
                    # 如果没有直接匹配，尝试从发件人名称中查找
                    if not outlet_info:
                        sender_name = msg.SenderName.lower()
                        for email, info in email_to_outlet.items():
                            if info['full_name'].lower() in sender_name:
                                outlet_info = info
                                outlet_key = info['short_name'].lower()  # 使用分店简称作为key
                                break
                
                # 检查是否是Weekly Order邮件
                subject = (msg.Subject or "").lower()
                is_weekly = "weekly" in subject or "order" in subject
                
                if not is_weekly:
                    continue
                
                # 只保留每个分店最新的邮件
                if outlet_key not in latest_messages or msg.ReceivedTime > latest_messages[outlet_key].ReceivedTime:
                    latest_messages[outlet_key] = msg
                    
            except Exception as e:
                print(f"Error processing message: {e}")
        
        return list(latest_messages.values())

    @classmethod
    def _download_attachments(cls, messages, save_path, email_to_outlet=None, week_no=None):
        """下载邮件附件，按分店重命名"""
        result = {
            "downloaded": 0,
            "skipped": 0,
            "matched_outlets": [],
            "unmatched_emails": set()
        }
        
        for msg in messages:
            try:
                sender_email = msg.SenderEmailAddress.lower()
                attachments = msg.Attachments
                
                if attachments.Count == 0:
                    continue
                
                # 尝试匹配分店
                outlet_info = None
                prefix = "UNKNOWN"
                
                if email_to_outlet:
                    outlet_info = email_to_outlet.get(sender_email)
                    
                    # 如果没有直接匹配，尝试从发件人名称中查找
                    if not outlet_info:
                        sender_name = msg.SenderName.lower()
                        for email, info in email_to_outlet.items():
                            if info['full_name'].lower() in sender_name:
                                outlet_info = info
                                break
                
                if outlet_info:
                    prefix = outlet_info['short_name']
                    result['matched_outlets'].append(prefix)
                else:
                    result['unmatched_emails'].add(sender_email)
                
                # 下载所有附件（Weekly Order通常只有一个附件）
                for att in attachments:
                    filename = att.FileName
                    
                    # 构建新文件名
                    if week_no:
                        new_filename = f"{prefix}_weekly_order_week{week_no}_{filename}"
                    else:
                        new_filename = f"{prefix}_{filename}"
                    
                    # 保存附件
                    full_path = os.path.join(save_path, new_filename)
                    att.SaveAsFile(full_path)
                    result['downloaded'] += 1
            
            except Exception as e:
                print(f"Error processing attachment: {e}")
                result['skipped'] += 1
        
        return result

# ========== Operation Supplies Monthly Order Core ==========
class OperationSuppliesOrder:
    """Operation Supplies Monthly Order Automation with MOQ calculation and amount display"""
    @staticmethod
    def get_monthly_order_data(master_file):
        """Extract outlet data and order quantities from master file"""
        try:
            wb = load_workbook(master_file, data_only=True)
            data_sheet = wb["Data"]
            
            # Extract outlet data
            outlets = []
            for row in data_sheet.iter_rows(min_row=2, max_col=7, values_only=True):
                if row[1] and row[2]:  # Ensure outlet and short name exist
                    outlets.append({
                        "brand": row[1],          # Column B - Brands
                        "outlet": row[2],         # Column C - Outlet (B2 value)
                        "short_name": row[3],     # Column D - Short Name (Sheet Name)
                        "full_name": row[4],      # Column E - Outlet Full Name (for F5)
                        "address": row[5],        # Column F - Address (for F6)
                        "delivery_day": row[6]    # Column G - Delivery Day
                    })
            
            # Extract order quantities and unit prices
            orders = defaultdict(dict)
            unit_prices = defaultdict(dict)
            
            # Get unit prices from supplier templates
            for supplier in ["Freshening", "Legacy", "Unikleen"]:
                if supplier in wb.sheetnames:
                    ws = wb[supplier]
                    
                    # Unit prices are in column D
                    if supplier == "Freshening":
                        # D12:D45 for 34 items
                        unit_prices[supplier] = [ws[f'D{i}'].value for i in range(12, 46)]
                    elif supplier == "Legacy":
                        # D12:D14 for 3 items
                        unit_prices[supplier] = [ws[f'D{i}'].value for i in range(12, 15)]
                    else:  # Unikleen
                        # D12:D29 for 18 items
                        unit_prices[supplier] = [ws[f'D{i}'].value for i in range(12, 30)]
            
            # Extract orders from outlet sheets
            for outlet in outlets:
                sheet_name = outlet["short_name"]
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    
                    # Freshening orders (L4:L37)
                    freshening = []
                    for row in range(4, 38):
                        cell_value = ws[f'L{row}'].value
                        freshening.append(cell_value if cell_value is not None else 0)
                    
                    # Legacy orders (L41:L43)
                    legacy = []
                    for row in range(41, 44):
                        cell_value = ws[f'L{row}'].value
                        legacy.append(cell_value if cell_value is not None else 0)
                    
                    # Unikleen orders (L47:L64)
                    unikleen = []
                    for row in range(47, 65):
                        cell_value = ws[f'L{row}'].value
                        unikleen.append(cell_value if cell_value is not None else 0)
                    
                    orders[sheet_name] = {
                        "freshening": freshening,
                        "legacy": legacy,
                        "unikleen": unikleen
                    }
            
            # Get supplier templates
            templates = {}
            for supplier in ["Freshening", "Legacy", "Unikleen"]:
                if supplier in wb.sheetnames:
                    templates[supplier] = wb[supplier]
            
            return outlets, orders, templates, unit_prices
        except Exception as e:
            return None, f"Error reading master file: {str(e)}"

    @classmethod
    def calculate_order_amounts(cls, orders, unit_prices):
        """Calculate order amounts for each outlet and supplier"""
        amounts = defaultdict(dict)
        
        for outlet, order_data in orders.items():
            for supplier in ["freshening", "legacy", "unikleen"]:
                items = order_data.get(supplier, [])
                prices = unit_prices.get(supplier.capitalize(), [])
                
                # Ensure we have prices for all items
                if len(prices) < len(items):
                    prices = prices + [0] * (len(items) - len(prices))
                
                # Calculate total amount
                total = sum(qty * price for qty, price in zip(items, prices) if price is not None)
                amounts[outlet][supplier] = total
        
        return amounts

    @classmethod
    def check_moq(cls, outlets, orders, unit_prices, log_callback=None):
        """Check MOQ requirements for each outlet and display order amounts"""
        results = {
            "freshening": defaultdict(list),
            "legacy": defaultdict(list),
            "unikleen": defaultdict(list)
        }
        
        # Calculate order amounts
        amounts = cls.calculate_order_amounts(orders, unit_prices)
        
        # Freshening MOQ check
        for outlet in outlets:
            short_name = outlet["short_name"]
            brand_type = outlet["brand"]
            amount = amounts.get(short_name, {}).get("freshening", 0)
            
            if brand_type == "Dine-In" and amount < 150:
                results["freshening"]["below_moq"].append(f"{short_name} (${amount:.2f} < $150)")
            elif brand_type == "GOGO" and amount < 100:
                results["freshening"]["below_moq"].append(f"{short_name} (${amount:.2f} < $100)")
            elif brand_type == "CNK" and amount < 150:
                results["freshening"]["below_moq"].append(f"{short_name} (${amount:.2f} < $150)")
            elif amount > 0:
                results["freshening"]["above_moq"].append(f"{short_name} (${amount:.2f})")
        
        # Legacy MOQ check
        for outlet in outlets:
            short_name = outlet["short_name"]
            quantities = orders.get(short_name, {}).get("legacy", [])
            
            total = sum(q for q in quantities if isinstance(q, (int, float)))
            cartons = total  # Assuming 1 carton per item
            
            if cartons < 2 and total > 0:
                results["legacy"]["below_moq"].append(f"{short_name} ({cartons} ctn < 2 ctn)")
            elif total > 0:
                results["legacy"]["above_moq"].append(f"{short_name} ({cartons} ctn)")
        
        # Unikleen MOQ check
        for outlet in outlets:
            short_name = outlet["short_name"]
            amount = amounts.get(short_name, {}).get("unikleen", 0)
            
            if amount < 80 and amount > 0:
                results["unikleen"]["below_moq"].append(f"{short_name} (${amount:.2f} < $80)")
            elif amount > 0:
                results["unikleen"]["above_moq"].append(f"{short_name} (${amount:.2f})")
        
        # Generate summary with amounts
        summary = "=== MOQ 檢查結果 (顯示訂購金額) ===\n"
        
        for supplier in ["freshening", "legacy", "unikleen"]:
            summary += f"\n** {supplier.capitalize()} **\n"
            
            if results[supplier]["below_moq"]:
                summary += "❌ 未達MOQ:\n"
                summary += "\n".join([f"  - {outlet}" for outlet in results[supplier]["below_moq"]]) + "\n"
            
            if results[supplier]["above_moq"]:
                summary += "✅ 已達MOQ:\n"
                summary += "\n".join([f"  - {outlet}" for outlet in results[supplier]["above_moq"]]) + "\n"
            
            if not results[supplier]["below_moq"] and not results[supplier]["above_moq"]:
                summary += "⚠️ 沒有訂單\n"
        
        if log_callback:
            log_callback(summary)
        
        return results, summary, amounts

    @classmethod
    def generate_supplier_files(cls, master_file, output_folder, outlets, orders, templates, amounts, log_callback=None):
        """Generate order files for each supplier using templates and display order amounts"""
        try:
            now = datetime.now()
            next_month = now.month + 1 if now.month < 12 else 1
            year = now.year if now.month < 12 else now.year + 1
            supplier_files = []
            
            # Create supplier files
            for supplier, template_ws in templates.items():
                # Create new workbook
                wb = Workbook()
                wb.remove(wb.active)  # Remove default sheet
                
                # Add outlets that have orders for this supplier
                outlet_count = 0
                for outlet in outlets:
                    short_name = outlet["short_name"]
                    outlet_orders = orders.get(short_name, {})
                    
                    if not outlet_orders:
                        continue
                    
                    # Get orders for this supplier
                    supplier_key = supplier.lower()
                    if supplier_key == "freshening":
                        order_items = outlet_orders.get("freshening", [])
                    elif supplier_key == "legacy":
                        order_items = outlet_orders.get("legacy", [])
                    else:  # Unikleen
                        order_items = outlet_orders.get("unikleen", [])
                    
                    # Get order amount
                    order_amount = amounts.get(short_name, {}).get(supplier_key, 0)
                    
                    # Only create sheet if there are orders
                    if any(qty > 0 for qty in order_items):
                        # Create new sheet for this outlet from template
                        new_ws = wb.create_sheet(title=short_name)
                        
                        # Copy template to new sheet
                        for row in template_ws.iter_rows():
                            for cell in row:
                                new_cell = new_ws.cell(
                                    row=cell.row, 
                                    column=cell.column, 
                                    value=cell.value
                                )
                                
                                # Copy styles
                                if cell.has_style:
                                    new_cell.font = copy.copy(cell.font)
                                    new_cell.border = copy.copy(cell.border)
                                    new_cell.fill = copy.copy(cell.fill)
                                    new_cell.number_format = cell.number_format
                                    new_cell.protection = copy.copy(cell.protection)
                                    new_cell.alignment = copy.copy(cell.alignment)
                        
                        # Copy column widths
                        for col_idx in range(1, template_ws.max_column + 1):
                            col_letter = get_column_letter(col_idx)
                            new_ws.column_dimensions[col_letter].width = template_ws.column_dimensions[col_letter].width
                        
                        # Copy row heights
                        for row_idx in range(1, template_ws.max_row + 1):
                            new_ws.row_dimensions[row_idx].height = template_ws.row_dimensions[row_idx].height
                        
                        # Copy merged cells
                        for merged_range in template_ws.merged_cells.ranges:
                            new_ws.merge_cells(str(merged_range))
                        
                        # Set outlet-specific values
                        new_ws['F5'] = outlet["full_name"]  # Outlet Full Name
                        new_ws['F6'] = outlet["address"]    # Address
                        
                        # Add delivery day for Freshening
                        if supplier == "Freshening":
                            new_ws['G5'] = outlet["delivery_day"]  # Delivery day
                        
                        # Add orders based on supplier type
                        if supplier == "Freshening":
                            # Freshening: L4:L37 (34 items)
                            for i, qty in enumerate(order_items[:34]):
                                new_ws[f'L{i+4}'] = qty
                            
                            # Display order amount in F46
                            new_ws['F46'] = order_amount
                            new_ws['F46'].number_format = '"$"#,##0.00'
                            
                        elif supplier == "Legacy":
                            # Legacy: L41:L43 (3 items)
                            for i, qty in enumerate(order_items[:3]):
                                new_ws[f'L{i+41}'] = qty
                            
                            # Display carton count in F16
                            cartons = sum(order_items[:3])
                            new_ws['F16'] = cartons
                            
                        else:  # Unikleen
                            # Unikleen: L47:L64 (18 items)
                            for i, qty in enumerate(order_items[:18]):
                                new_ws[f'L{i+47}'] = qty
                            
                            # Display order amount in F31
                            new_ws['F31'] = order_amount
                            new_ws['F31'].number_format = '"$"#,##0.00'
                        
                        outlet_count += 1
                
                # Save file only if it has sheets
                if outlet_count > 0:
                    file_name = f"{supplier}_Order_{year}_{next_month:02d}.xlsx"
                    file_path = os.path.join(output_folder, file_name)
                    wb.save(file_path)
                    supplier_files.append(file_name)
                    
                    if log_callback:
                        log_callback(f"✅ 已保存供應商文件: {file_path} (包含 {outlet_count} 個分店)")
                else:
                    if log_callback:
                        log_callback(f"⚠️ {supplier} 沒有訂單，未生成文件")
            
            return True, supplier_files
        except Exception as e:
            return False, f"生成供應商文件時出錯: {str(e)}\n{traceback.format_exc()}"

    @classmethod
    def process_order(cls, master_file, output_folder, log_callback=None, progress_callback=None):
        """Main processing function for operation supplies with MOQ and amounts"""
        try:
            # Step 1: Read data from master file
            if log_callback:
                log_callback(f"讀取主文件: {os.path.basename(master_file)}")
            outlets, orders, templates, unit_prices = cls.get_monthly_order_data(master_file)
            
            if not outlets:
                return False, "無法讀取分店數據，請檢查Data工作表"
            
            # Step 2: Calculate order amounts and check MOQ requirements
            if log_callback:
                log_callback("\n計算訂購金額並檢查MOQ要求...")
            moq_results, moq_summary, amounts = cls.check_moq(outlets, orders, unit_prices, log_callback)
            
            # Step 3: Generate supplier files with amounts displayed
            if log_callback:
                log_callback("\n生成供應商訂單文件並顯示訂購金額...")
            success, supplier_files = cls.generate_supplier_files(
                master_file, output_folder, outlets, orders, templates, amounts, log_callback
            )
            
            if not success:
                return False, supplier_files
            
            # Step 4: Create result summary
            result = (
                f"=== 營運用品月訂單處理完成 ===\n\n"
                f"📊 MOQ 檢查結果:\n{moq_summary}\n\n"
                f"📁 生成的供應商文件:\n" + 
                "\n".join([f"  - {file}" for file in supplier_files])
            )
            
            return True, result
        except Exception as e:
            return False, f"處理過程中出錯: {str(e)}\n{traceback.format_exc()}"


class SushiExpressApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"Sushi Express Automation Tool v{VERSION}")
        self.geometry("1200x900")
        self.minsize(1000,800)
        self.configure(fg_color=DARK_BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Placeholders
        self.download_folder_var = None
        self.checklist_folder_var = None
        self.folder_vars = {}
        self.master_file_var = None
        self.output_folder_var = None
        self.progress_popup = None
        self.mapping_popup = None
        self.outlet_config_var = None
        self.delivery_config_var = None

        self._setup_ui()
        self.show_login()

    def _setup_ui(self):
        # Header and content layout
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        self.header_frame = ctk.CTkFrame(self.main_container, fg_color=DARK_PANEL, corner_radius=24, height=150)
        self.header_frame.pack(fill="x", padx=10, pady=10)
        # Center logo
        self.logo_img = load_image(LOGO_PATH)
        if self.logo_img:
            ctk.CTkLabel(self.header_frame, image=self.logo_img, text="").pack(expand=True)
        # Title labels
        self.title_label = ctk.CTkLabel(self.header_frame, text="", font=FONT_TITLE, text_color=ACCENT_BLUE)
        self.subtitle_label = ctk.CTkLabel(self.header_frame, text="", font=FONT_SUB, text_color=TEXT_COLOR)
        self.title_label.pack()
        self.subtitle_label.pack()

        # Buttons frame
        self.buttons_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.buttons_frame.pack(fill="x", padx=10, pady=10)

        # Content frame
        self.content_frame = ctk.CTkFrame(self.main_container, fg_color=DARK_PANEL, corner_radius=24)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.content_frame.pack_propagate(False)

    def _on_close(self):
        if messagebox.askyesno(t("exit_system"), t("exit_confirm")):
            self.destroy()

    def show_login(self):
        for w in self.content_frame.winfo_children(): w.destroy()
        self.login_frame = ctk.CTkFrame(self.content_frame, fg_color=DARK_PANEL, corner_radius=24, width=450, height=400)
        self.login_frame.place(relx=0.5, rely=0.5, anchor="center")
        self.pwd_entry = ctk.CTkEntry(self.login_frame, show="*", font=FONT_SUB, width=300, placeholder_text=t("password"), fg_color=ENTRY_BG, height=45)
        self.pwd_entry.pack(pady=(20,10))
        self.pwd_entry.bind("<Return>", lambda e: self._try_login())
        ctk.CTkButton(self.login_frame, text=t("login_btn"), command=self._try_login, width=200).pack(pady=(0,20))
        ctk.CTkLabel(self.login_frame, text=f"Version {VERSION} | {DEVELOPER}", font=("Consolas",12), text_color="#64748B").pack(side="bottom", pady=10)

    def _try_login(self):
        if self.pwd_entry.get() == PASSWORD:
            self.login_frame.destroy()
            self.show_main_menu()
        else:
            messagebox.showerror(t("error"), t("incorrect_pw"))

    def show_main_menu(self):
        for w in self.content_frame.winfo_children(): w.destroy()
        self.title_label.configure(text=t("main_title"))
        self.subtitle_label.configure(text=t("select_function"))
        # Clear old buttons
        for w in self.buttons_frame.winfo_children(): w.destroy()
        # Create buttons
        funcs = [ (t("download_title"), self.show_download_ui),
                  (t("checklist_title"), self.show_checklist_ui),
                  (t("automation_title"), self.show_automation_ui),
                  ("營運用品月訂單", self.show_operation_supplies_ui),
                  (t("exit_system"), self._on_close) ]
        for text, cmd in funcs:
            btn = ctk.CTkButton(self.buttons_frame, text=text, command=cmd)
            btn.pack(side="left", padx=10, pady=10, expand=True)

    def show_function_ui(self, title, subtitle, content_callback):
        for w in self.content_frame.winfo_children(): w.destroy()
        self.title_label.configure(text=title)
        self.subtitle_label.configure(text=subtitle)
        content = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)
        content_callback(content)
        ctk.CTkButton(self.content_frame, text=t("back_to_menu"), command=self.show_main_menu).pack(pady=10)

    def show_download_ui(self):
        def build(c):
            self.download_folder_var = ctk.StringVar()
            self.config_file_var = ctk.StringVar()  # 新增配置文件变量
            
            # 原有文件夹选择
            row1 = ctk.CTkFrame(c)
            row1.pack(fill="x", pady=5)
            ctk.CTkLabel(row1, text="下载文件夹:").pack(side="left")
            ctk.CTkEntry(row1, textvariable=self.download_folder_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=t("browse"), command=self._select_download_folder).pack(side="left")
            
            # 新增配置文件选择
            row2 = ctk.CTkFrame(c)
            row2.pack(fill="x", pady=5)
            ctk.CTkLabel(row2, text="分店配置文件 (可选):").pack(side="left")
            ctk.CTkEntry(row2, textvariable=self.config_file_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=t("browse"), command=self._select_config_file).pack(side="left")
            
            ctk.CTkButton(c, text=t("start_download"), command=self._run_download).pack(pady=20)
        
        self.show_function_ui(t("download_title"), t("download_desc"), build)

    def _select_download_folder(self):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.download_folder_var.set(f)
    
    def _select_config_file(self):
        f = filedialog.askopenfilename(
            title="选择分店配置文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if f: 
            self.config_file_var.set(f)

    def _run_download(self):
        folder = self.download_folder_var.get()
        if not folder:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        
        config_file = self.config_file_var.get() or None
        OutlookDownloader.download_weekly_orders(
            folder, 
            config_file=config_file,
            callback=lambda msg: messagebox.showinfo(t("download_summary"), msg)
        )

    def show_checklist_ui(self):
        def build(c):
            self.checklist_folder_var = ctk.StringVar()
            self.delivery_config_var = ctk.StringVar()  # 新增供应商配置
            
            # 原有文件夹选择
            row1 = ctk.CTkFrame(c)
            row1.pack(fill="x", pady=5)
            ctk.CTkLabel(row1, text="订单文件夹:").pack(side="left")
            ctk.CTkEntry(row1, textvariable=self.checklist_folder_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=t("browse"), command=self._select_checklist_folder).pack(side="left")
            
            # 新增供应商配置选择
            row2 = ctk.CTkFrame(c)
            row2.pack(fill="x", pady=5)
            ctk.CTkLabel(row2, text="送货日期配置 (可选):").pack(side="left")
            ctk.CTkEntry(row2, textvariable=self.delivery_config_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=t("browse"), command=self._select_delivery_config).pack(side="left")
            
            ctk.CTkButton(c, text=t("run_check"), command=self._run_checklist).pack(pady=20)

    def _select_checklist_folder(self):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.checklist_folder_var.set(f)
    
    def _select_delivery_config(self):
        f = filedialog.askopenfilename(
            title="选择送货日期配置",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if f: 
            self.delivery_config_var.set(f)

    def _run_checklist(self):
        folder = self.checklist_folder_var.get()
        if not folder:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        
        config_file = self.delivery_config_var.get() or None
        result = OrderChecker.run_checklist(folder, config_file)
        ScrollableMessageBox(self, t("check_results"), result)

    def show_automation_ui(self):
        def build(c):
            self.folder_vars = {}
            self.outlet_config_var = ctk.StringVar()
            self.delivery_config_var = ctk.StringVar()
            
            # 原有文件夹选择
            for label, key in [(t("source_folder"),"source"),(t("supplier_folder"),"supplier"),(t("outlet_folder"),"outlet")]:
                row = ctk.CTkFrame(c); row.pack(pady=5)
                var = ctk.StringVar(); self.folder_vars[key] = var
                ctk.CTkLabel(row, text=label).pack(side="left", padx=5)
                ctk.CTkEntry(row, textvariable=var, state="readonly").pack(side="left", padx=5)
                ctk.CTkButton(row, text=t("browse"), command=lambda k=key: self._select_folder(k)).pack(side="left", padx=5)
            
            # 新增分店配置选择
            row = ctk.CTkFrame(c); row.pack(pady=5)
            ctk.CTkLabel(row, text="分店配置文件 (可选):").pack(side="left", padx=5)
            ctk.CTkEntry(row, textvariable=self.outlet_config_var, state="readonly").pack(side="left", expand=True, padx=5)
            ctk.CTkButton(row, text=t("browse"), command=lambda: self._select_config(self.outlet_config_var)).pack(side="left", padx=5)
            
            # 新增送货日期配置选择
            row = ctk.CTkFrame(c); row.pack(pady=5)
            ctk.CTkLabel(row, text="送货日期配置 (可选):").pack(side="left", padx=5)
            ctk.CTkEntry(row, textvariable=self.delivery_config_var, state="readonly").pack(side="left", expand=True, padx=5)
            ctk.CTkButton(row, text=t("browse"), command=lambda: self._select_config(self.delivery_config_var)).pack(side="left", padx=5)
            
            ctk.CTkButton(c, text=t("start_automation"), command=self._run_automation).pack(pady=20)

    def _select_folder(self, key):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.folder_vars[key].set(f)
    
    def _select_config(self, var):
        f = filedialog.askopenfilename(
            title="选择配置文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if f: 
            var.set(f)

    def _run_automation(self):
        src, sup, out = self.folder_vars.get('source'), self.folder_vars.get('supplier'), self.folder_vars.get('outlet')
        if not all([src.get(), sup.get(), out.get()]): 
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        
        # 读取配置
        outlet_config = None
        if self.outlet_config_var.get():
            try:
                outlet_config = OutlookDownloader.read_outlet_config(self.outlet_config_var.get())
            except Exception as e:
                messagebox.showerror("错误", f"无法读取分店配置: {str(e)}")
                return
        
        delivery_config = self.delivery_config_var.get() or None
        
        p = ProgressPopup(self, t("processing_orders"))
        self.progress_popup = p
        def logcb(m): p.log(m)
        def mapcb(m): 
            mp = MappingPopup(self, t("outlet_suppliers")); mp.update_mapping(m)
            self.mapping_popup = mp
        
        threading.Thread(target=lambda: self._thread_task(
            lambda: OrderAutomation.run_automation(
                src.get(), sup.get(), out.get(), 
                outlet_config, delivery_config, logcb, mapcb)
        )).start()

    def show_operation_supplies_ui(self):
        def build(c):
            self.master_file_var = ctk.StringVar(); self.output_folder_var = ctk.StringVar()
            for label,var,fn in [("Select Master File", self.master_file_var, self._select_master_file),("Select Output Folder", self.output_folder_var, self._select_output_folder)]:
                row = ctk.CTkFrame(c); row.pack(pady=5)
                ctk.CTkLabel(row, text=label).pack(side="left", padx=5)
                ctk.CTkEntry(row, textvariable=var, state="readonly").pack(side="left", padx=5)
                ctk.CTkButton(row, text=t("browse"), command=fn).pack(side="left", padx=5)
            ctk.CTkButton(c, text="Start Processing", command=self._run_operation_supplies).pack(pady=20)
        self.show_function_ui("Operation Supplies", "", build)

    def _select_master_file(self):
        f = filedialog.askopenfilename(title=t("select_folder"), filetypes=[("Excel","*.xlsx;*.xls")])
        if f: self.master_file_var.set(f)

    def _select_output_folder(self):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.output_folder_var.set(f)

    def _run_operation_supplies(self):
        mf, of = self.master_file_var.get(), self.output_folder_var.get()
        if not mf or not of: messagebox.showwarning(t("warning"), t("folder_warning")); return
        p = ProgressPopup(self, t("processing"))
        self.progress_popup = p
        def logcb(m): p.log(m)
        threading.Thread(target=lambda: self._thread_task(lambda: OperationSuppliesOrder.process_order(mf, of, logcb))).start()

    def _thread_task(self, fn):
        success, msg = fn()
        self.after(0, lambda: messagebox.showinfo(t("success") if success else t("warning"), msg))

# Entry point
if __name__ == '__main__':
    try:
        app = SushiExpressApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"Startup failed: {e}")
        sys.exit(1)
l
