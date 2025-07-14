import os
import sys
import threading
import time
import re
import copy
import csv
import traceback
from datetime import datetime, timedelta
from collections import defaultdict
from dateutil.parser import parse
import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import codecs
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from openpyxl.cell.cell import MergedCell
import glob
import xlwings as xw

# ======== 资源路径处理 ========
def resource_path(relative_path):
    """获取资源文件的绝对路径"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ========== 全局配置 ==========
LOGO_PATH = resource_path("SELOGO22 - 01.png")
PASSWORD = "OPS123"
VERSION = "6.6.0"
DEVELOPER = "OPS - Voon Kee"

# 主题配置
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# 颜色定义
DARK_BG = "#0f172a"
DARK_PANEL = "#1e293b"
ACCENT_BLUE = "#3b82f6"
BTN_HOVER = "#60a5fa"
ACCENT_GREEN = "#10b981"
ACCENT_RED = "#ef4444"
ACCENT_PURPLE = "#8b5cf6"
ENTRY_BG = "#334155"
TEXT_COLOR = "#e2e8f0"
PANEL_BG = "#1e293b"
TEXTBOX_BG = "#0f172a"

# 字体配置
FONT_TITLE = ("Microsoft JhengHei", 24, "bold")
FONT_BIGBTN = ("Microsoft JhengHei", 16, "bold")
FONT_MID = ("Microsoft JhengHei", 14)
FONT_SUB = ("Microsoft JhengHei", 12)
FONT_ZH = ("Microsoft JhengHei", 12)
FONT_EN = ("Segoe UI", 11, "italic")
FONT_LOG = ("Consolas", 14)

# ========== 多语言支持 ==========
def t(text):
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
        "automation_title": "訂單自動整合\nOrder Automation",
        "automation_desc": "請選擇三個必要的資料夾\nSelect required folders",
        "source_folder": "來源資料夾 (Weekly Orders)\nSource Folder (Weekly Orders)",
        "supplier_folder": "供應商資料夾 (Supplier)\nSupplier Folder",
        "outlet_folder": "分店資料夾 (Outlet)\nOutlet Folder",
        "start_automation": "開始整合檔案\nStart Automation",
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
        "send_emails": "發送郵件\nSend Emails",
        "operation_supplies": "營運用品\nOperation Supplies",
    }
    return translations.get(text, text)

def get_contrast_color(bg_color):
    # 簡單亮色/暗色對比
    if isinstance(bg_color, str) and bg_color.startswith('#'):
        r = int(bg_color[1:3], 16)
        g = int(bg_color[3:5], 16)
        b = int(bg_color[5:7], 16)
        luminance = (0.299*r + 0.587*g + 0.114*b)
        return '#000000' if luminance > 186 else '#ffffff'
    return '#ffffff'

# ========== 工具函数 ==========
def load_image(path, max_size=(400, 130)):
    """安全加载图像"""
    try:
        if not os.path.exists(path):
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

# ========== 自定义UI组件 ==========
class GlowButton(ctk.CTkButton):
    """发光效果按钮（美化版）"""
    def __init__(self, master, text=None, glow_color=ACCENT_BLUE, **kwargs):
        super().__init__(master, text=text, **kwargs)
        self._glow_color = glow_color
        self._setup_style()
        self._bind_events()

    def _setup_style(self):
        self.configure(
            border_width=0,
            fg_color=self._glow_color,
            hover_color=self._adjust_color(self._glow_color, 40),
            text_color=get_contrast_color(self._glow_color),
            corner_radius=22,  # 更圆角
            font=("Microsoft JhengHei", 20, "bold"),  # 更大更粗
            height=70,  # 更高
            width=340,  # 更宽
            anchor="center"
        )

    def _bind_events(self):
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)

    def _on_enter(self, event=None):
        self.configure(border_width=4, border_color=self._adjust_color(self._glow_color, 60), fg_color=self._adjust_color(self._glow_color, 20))

    def _on_leave(self, event=None):
        self.configure(border_width=0, fg_color=self._glow_color)
    
    @staticmethod
    def _adjust_color(color, amount):
        if isinstance(color, tuple) and len(color) >= 3:
            r, g, b = color[:3]
            adjusted = tuple(min(255, max(0, x + amount)) for x in (r, g, b))
            if len(color) == 4:
                return adjusted + (color[3],)
            return adjusted
        else:
            color = color.lstrip('#')
            rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            adjusted = tuple(min(255, max(0, x + amount)) for x in rgb)
            return f"#{adjusted[0]:02x}{adjusted[1]:02x}{adjusted[2]:02x}"

class ProgressPopup(ctk.CTkToplevel):
    """进度显示弹窗"""
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self.parent = parent
        self.configure(fg_color=DARK_BG)
        
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
        
        close_btn = GlowButton(
            self,
            text=t("close"),
            command=self.destroy_popup,
            width=120,
            height=40
        )
        close_btn.pack(pady=10)
    
    def destroy_popup(self):
        self.destroy()
        self.parent.progress_popup = None
    
    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", message)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

class MappingPopup(ctk.CTkToplevel):
    """分店-供应商映射显示"""
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self.parent = parent
        self.configure(fg_color=DARK_BG)
        
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
        
        close_btn = GlowButton(
            self,
            text=t("close"),
            command=self.destroy_popup,
            width=120,
            height=40
        )
        close_btn.pack(pady=10)
    
    def destroy_popup(self):
        self.destroy()
        self.parent.mapping_popup = None
    
    def update_mapping(self, mapping):
        self.mapping_text.configure(state="normal")
        self.mapping_text.delete("1.0", "end")
        self.mapping_text.insert("1.0", mapping)
        self.mapping_text.configure(state="disabled")

class ScrollableMessageBox(ctk.CTkToplevel):
    """可滚动消息框"""
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.geometry("900x700")
        self.transient(parent)
        self.grab_set()
        self.configure(fg_color=DARK_BG)
        
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
        
        close_btn = GlowButton(
            self,
            text=t("close"),
            command=self.destroy,
            width=120,
            height=40
        )
        close_btn.pack(pady=10)

class EmailConfirmationDialog(ctk.CTkToplevel):
    """邮件发送确认对话框（只顯示純文字，保留空行）"""
    def __init__(self, parent, mail_item, supplier_name, outlet_name, attachment_path, on_confirm):
        super().__init__(parent)
        self.title(f"确认邮件 - {supplier_name}")
        self.geometry("800x700")
        self.transient(parent)
        self.grab_set()
        self.mail_item = mail_item
        self.on_confirm = on_confirm
        self.attachment_path = attachment_path
        self.configure(fg_color=DARK_BG)
        info_frame = ctk.CTkFrame(self, fg_color=DARK_PANEL, corner_radius=12)
        info_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkLabel(info_frame, text=f"收件人(To)/To: {mail_item.To}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"抄送(CC)/CC: {mail_item.CC}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"主题/Subject: {mail_item.Subject}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"供应商/Supplier: {supplier_name}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"分店/Outlet: {outlet_name}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        if attachment_path:
            file_name = os.path.basename(attachment_path)
            ctk.CTkLabel(info_frame, text=f"附件/Attachment: {file_name}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        body_frame = ctk.CTkFrame(self, fg_color=DARK_PANEL, corner_radius=12)
        body_frame.pack(fill="both", expand=True, padx=20, pady=10)
        ctk.CTkLabel(body_frame, text="邮件正文/Email Body:", font=FONT_BIGBTN).pack(anchor="w", padx=10, pady=5)
        # 只顯示純文字，保留空行，避免名字被拆行
        import html
        raw_html = mail_item.HTMLBody or ""
        import re
        # 1. 只把結束分行標籤換成換行
        text = re.sub(r'(<br\s*/?>|</div>|</p>|</li>|</tr>|</td>)', '\n', raw_html, flags=re.IGNORECASE)
        # 2. 把開始標籤直接移除（不換行）
        text = re.sub(r'(<div>|<p>|<li>|<tr>|<td>)', '', text, flags=re.IGNORECASE)
        # 3. 移除其他 HTML 標籤
        text = re.sub(r'<[^>]+>', '', text)
        # 4. decode HTML entity
        text = html.unescape(text)
        # 5. 移除開頭/結尾多餘空白
        text = text.strip()
        self.body_text = ctk.CTkTextbox(body_frame, wrap="word", height=300, font=FONT_MID)
        self.body_text.pack(fill="both", expand=True, padx=10, pady=5)
        self.body_text.insert("1.0", text)
        self.body_text.configure(state="disabled")
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(
            btn_frame, 
            text="发送邮件/Send Email", 
            command=self._send_email,
            fg_color=ACCENT_GREEN,
            hover_color=BTN_HOVER
        ).pack(side="right", padx=10)
        ctk.CTkButton(
            btn_frame, 
            text="取消/Cancel", 
            command=self.destroy
        ).pack(side="right", padx=10)
        ctk.CTkButton(
            btn_frame, 
            text="编辑正文/Edit Body", 
            command=self._edit_body,
            fg_color=ACCENT_BLUE,
            hover_color=BTN_HOVER
        ).pack(side="left", padx=10)
    def _edit_body(self):
        self.body_text.configure(state="normal")
    def _send_email(self):
        try:
            # 取得內容並轉成 HTML 格式
            body = self.body_text.get("1.0", "end-1c")
            html_body = body.replace("\n", "<br>")
            self.mail_item.HTMLBody = html_body
            self.mail_item.Send()
            self.on_confirm(True, "邮件发送成功！")
            self.destroy()
        except Exception as e:
            import traceback
            print("Send email error:", e)
            print(traceback.format_exc())
            self.on_confirm(False, f"邮件发送失败: {str(e)}")
            self.destroy()

class NavigationButton(ctk.CTkButton):
    """自定义导航按钮，支持选中状态"""
    def __init__(self, master, text, command, **kwargs):
        super().__init__(master, text=text, command=command, **kwargs)
        self.default_color = DARK_PANEL
        self.selected_color = ACCENT_BLUE
        self.hover_color = BTN_HOVER
        self.is_selected = False
        self.default_border_color = ACCENT_BLUE
        self.selected_border_color = ACCENT_GREEN
        self.configure(
            fg_color=self.default_color,
            hover_color=self.hover_color,
            text_color=get_contrast_color(self.default_color),
            border_width=2,
            border_color=self.default_border_color,
            corner_radius=10,
            font=FONT_BIGBTN,
            height=50,
            anchor="center"
        )
        self._bind_events()
    def _bind_events(self):
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
    def _on_enter(self, event=None):
        if not self.is_selected:
            self.configure(fg_color=self._adjust_color(self.default_color, 20))
    def _on_leave(self, event=None):
        if not self.is_selected:
            self.configure(fg_color=self.default_color)
    def select(self):
        self.is_selected = True
        self.configure(
            fg_color=self.selected_color,
            border_width=2,
            border_color=self.selected_border_color,
            text_color=get_contrast_color(self.selected_color)
        )
    def deselect(self):
        self.is_selected = False
        self.configure(
            fg_color=self.default_color,
            border_width=2,
            border_color=self.default_border_color,
            text_color=get_contrast_color(self.default_color)
        )
    @staticmethod
    def _adjust_color(color, amount):
        if isinstance(color, str) and color.startswith("#"):
            color = color.lstrip('#')
            rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            adjusted = tuple(min(255, max(0, x + amount)) for x in rgb)
            return f"#{adjusted[0]:02x}{adjusted[1]:02x}{adjusted[2]:02x}"
        return color

# ========== 送货日期验证工具 ==========
class DeliveryDateValidator:
    """送货日期验证工具"""
    
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
        """加载送货日期配置"""
        try:
            with open(config_file, mode='r', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    supplier = row['supplier'].strip().upper()
                    outlet = row['outlet_code'].strip().upper()
                    days = self.parse_delivery_days(row['delivery_days'])
                    
                    if outlet == "ALL":
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
        supplier = supplier.strip().upper()
        outlet_code = outlet_code.strip().upper()
        outlet_specific = self.schedule.get(supplier, {}).get(outlet_code)
        if outlet_specific is not None:
            return outlet_specific
        return self.schedule.get(supplier, {}).get('*', set())

    def validate_order(self, supplier, outlet_code, order_date, log_callback=None):
        supplier = supplier.strip().upper()
        outlet_code = outlet_code.strip().upper()
        delivery_days = self.get_delivery_days(supplier, outlet_code)
        
        if not delivery_days:
            return True
        
        try:
            if isinstance(order_date, (int, float)):
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

# ========== 统一配置管理器 ==========
class UnifiedConfigManager:
    """管理统一的Excel配置文件"""
    
    def __init__(self, config_path=None):
        self.config_path = config_path
        self.outlets = []
        self.suppliers = []
        self.delivery_schedule = []
        self.email_templates = {}
        self.supplier_requirements = {}
        
        if config_path:
            self.load_config(config_path)
    
    def load_config(self, config_path):
        """从Excel文件加载配置"""
        try:
            wb = load_workbook(config_path, data_only=True)
            
            # 读取分店信息
            if "Outlets" in wb.sheetnames:
                ws = wb["Outlets"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        self.outlets.append({
                            "code": row[0].strip().upper() if isinstance(row[0], str) else str(row[0]).strip().upper(),
                            "name": row[1].strip() if len(row) > 1 and isinstance(row[1], str) else str(row[1]).strip() if len(row) > 1 and row[1] is not None else "",
                            "email": row[2].strip().lower() if len(row) > 2 and isinstance(row[2], str) else str(row[2]).strip().lower() if len(row) > 2 and row[2] is not None else "",
                            "address": row[3].strip() if len(row) > 3 and isinstance(row[3], str) else str(row[3]).strip() if len(row) > 3 and row[3] is not None else "",
                            "delivery_day": row[4].strip() if len(row) > 4 and isinstance(row[4], str) else str(row[4]).strip() if len(row) > 4 and row[4] is not None else "",
                            "brand": row[5].strip() if len(row) > 5 and isinstance(row[5], str) else str(row[5]).strip() if len(row) > 5 and row[5] is not None else "Dine-In"
                        })
            
            # 读取供应商信息
            if "Suppliers" in wb.sheetnames:
                ws = wb["Suppliers"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        self.suppliers.append({
                            "name": row[0].strip() if isinstance(row[0], str) else str(row[0]).strip(),
                            "email": row[1].strip().lower() if len(row) > 1 and isinstance(row[1], str) else str(row[1]).strip().lower() if len(row) > 1 and row[1] is not None else "",
                            "contact": row[2].strip() if len(row) > 2 and isinstance(row[2], str) else str(row[2]).strip() if len(row) > 2 and row[2] is not None else "",
                            "phone": row[3].strip() if len(row) > 3 and isinstance(row[3], str) else str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
                        })
            
            # 读取配送日程
            if "Delivery Schedule" in wb.sheetnames:
                ws = wb["Delivery Schedule"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        self.delivery_schedule.append({
                            "supplier": row[0].strip() if isinstance(row[0], str) else str(row[0]).strip(),
                            "outlet_code": row[1].strip().upper() if len(row) > 1 and isinstance(row[1], str) else str(row[1]).strip().upper() if len(row) > 1 and row[1] is not None else "ALL",
                            "delivery_days": row[2].strip() if len(row) > 2 and isinstance(row[2], str) else str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                        })
            
            # 读取邮件模板
            if "Email Templates" in wb.sheetnames:
                ws = wb["Email Templates"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and len(row) > 0 and row[0]:
                        self.email_templates[row[0].strip() if isinstance(row[0], str) else str(row[0]).strip()] = {
                            "subject": row[1].strip() if len(row) > 1 and isinstance(row[1], str) else str(row[1]).strip() if len(row) > 1 and row[1] is not None else "",
                            "body": row[2].strip() if len(row) > 2 and isinstance(row[2], str) else str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                        }
            
            # 读取供应商要求
            if "Supplier Requirements" in wb.sheetnames:
                ws = wb["Supplier Requirements"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        supplier_name = row[0].strip() if isinstance(row[0], str) else str(row[0]).strip()
                        outlet_codes = [code.strip().upper() for code in str(row[1]).split(",") if code and hasattr(code, 'strip') and code.strip()] if len(row) > 1 and row[1] is not None else []
                        if supplier_name not in self.supplier_requirements:
                            self.supplier_requirements[supplier_name] = []
                        self.supplier_requirements[supplier_name] = outlet_codes
            
            return True, f"成功加载配置文件: {len(self.outlets)} 分店, {len(self.suppliers)} 供应商"
        except Exception as e:
            return False, f"加载配置文件失败: {str(e)}"
    
    def get_outlet(self, code):
        # 根據分店代碼獲取分店全名
        if not code:
            return None
        if isinstance(code, str):
            code = code.strip().upper()
        else:
            code = str(code).upper()
        for outlet in self.outlets:
            if outlet['short_name'].upper() == code:
                return outlet['full_name']
        return None
    
    def get_supplier(self, name):
        # 根據供應商名稱獲取標準名稱
        if not name:
            return None
        if isinstance(name, str):
            name = name.strip().upper()
        else:
            name = str(name).upper()
        for supplier in self.suppliers:
            if supplier['name'].upper() == name:
                return supplier['name']
        return None
    
    def get_delivery_schedule(self, supplier, outlet_code):
        """获取特定供应商-分店的配送日程"""
        for schedule in self.delivery_schedule:
            if schedule["supplier"] == supplier and schedule["outlet_code"] == outlet_code:
                return schedule["delivery_days"]
        
        for schedule in self.delivery_schedule:
            if schedule["supplier"] == supplier and schedule["outlet_code"] == "ALL":
                return schedule["delivery_days"]
        
        return None
    
    def get_required_outlets(self, supplier_name):
        """获取供应商必须包含的分店列表"""
        return self.supplier_requirements.get(supplier_name, [])

# ========== 邮件发送管理器 ==========
class EmailSender:
    """处理邮件发送功能"""
    
    def __init__(self, config_manager=None):
        self.config_manager = config_manager
    
    def _get_standard_subject(self, supplier_name):
        """生成标准邮件主题"""
        now = datetime.now()
        month_name = now.strftime("%B")
        week_in_month = (now.day - 1) // 7 + 1
        return f"Sushi Express Weekly Order - {supplier_name} - {month_name} - Week {week_in_month}"
    
    def get_to_cc_emails(self, supplier_name, config_path):
        """根據 config excel 取得 TO/CC 郵件，並自動加 opsadmin 及 purchasing.admin，且CC不重複"""
        wb = load_workbook(config_path, data_only=True)
        to_emails = []
        cc_emails = []
        if "Suppliers" in wb.sheetnames:
            ws = wb["Suppliers"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[0]:
                    continue
                if str(row[0]).strip() == supplier_name:
                    typ = str(row[1]).strip().upper() if len(row) > 1 and row[1] else "TO"
                    email = str(row[2]).strip() if len(row) > 2 and row[2] else ""
                    if not email:
                        continue
                    if typ == "TO":
                        to_emails.append(email)
                    elif typ == "CC":
                        cc_emails.append(email)
        # 一定要加 opsadmin 及 purchasing.admin
        cc_emails.append("Opsadmin@sushiexpress.com.sg")
        cc_emails.append("Purchasing.admin@sushiexpress.com.sg")
        # 去重（不分大小寫）
        seen = set()
        cc_unique = []
        for e in cc_emails:
            elower = e.lower()
            if elower not in seen:
                cc_unique.append(e)
                seen.add(elower)
        return ",".join(to_emails), ",".join(cc_unique)
    
    def send_email(self, to_emails, cc_emails, supplier_name, body, attachment_path=None, account_idx=None):
        try:
            import win32com.client
            import os
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            # 指定帳號發送
            if account_idx is not None:
                mapi = outlook.GetNamespace("MAPI")
                accounts = [mapi.Folders.Item(i + 1) for i in range(mapi.Folders.Count)]
                if 0 <= account_idx < len(accounts):
                    account = accounts[account_idx]
                    if hasattr(mail, 'SendUsingAccount'):
                        for acc in outlook.Session.Accounts:
                            if acc.DisplayName == account.Name:
                                mail.SendUsingAccount = acc
                                break
            # 處理收件人
            to_emails_fixed = to_emails.replace('，', ';').replace(',', ';') if to_emails else ''
            cc_emails_fixed = cc_emails.replace('，', ';').replace(',', ';') if cc_emails else ''
            mail.To = to_emails_fixed
            mail.CC = cc_emails_fixed
            mail.Subject = self._get_standard_subject(supplier_name)

            # 保留所有空行，建議用者在 body 內容用雙換行 (\n\n) 來產生明顯空行
            mail.Body = body.replace('\n', '\n\n')

            # 附件
            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            # 已移除 gif 相關功能

            # 改用 HTMLBody 發送，保證格式
            html_body = body
            if not any(tag in body.lower() for tag in ['<br', '<p', '<div', '<table', '<ul', '<ol', '<li', '<b', '<strong', '<em', '<span']):
                html_body = body.replace('\n', '<br>')
            mail.HTMLBody = html_body

            return mail
        except Exception as e:
            import traceback
            print(traceback.format_exc())
            return None, f"创建邮件失败: {str(e)}"

# ========== 订单自动化核心 ==========
class OrderAutomation:
    """订单自动化工具"""
    
    def __init__(self, outlet_config=None):
        # outlet_config: list of dicts with keys 'short_name', 'full_name'
        self.outlet_name_map = {}
        if outlet_config:
            for o in outlet_config:
                full = o.get('full_name', '').strip().lower()
                short = o.get('short_name', '').strip().upper()
                if full:
                    self.outlet_name_map[full] = short
                if short:
                    self.outlet_name_map[short.lower()] = short
    def get_short_code(self, f5val):
        val = (str(f5val) or '').strip().lower()
        if val in self.outlet_name_map:
            return self.outlet_name_map[val]
        # 雙向模糊比對
        for full, short in self.outlet_name_map.items():
            if full in val or val in full:
                return short
            # 單字比對
            for word in val.split():
                if word and word in full:
                    return short
        return 'UNKNOWN'

    @staticmethod
    def is_valid_date(cell_value, next_week_start, next_week_end):
        """检查是否为有效日期"""
        try:
            if isinstance(cell_value, (int, float)):
                base_date = datetime(1899, 12, 30)
                parsed = base_date + timedelta(days=cell_value)
            else:
                parsed = parse(str(cell_value), fuzzy=True, dayfirst=False)
            return next_week_start.date() <= parsed.date() <= next_week_end.date()
        except:
            return False

    @classmethod
    def find_delivery_date_row(cls, ws, next_week_start, next_week_end, max_rows=150, file_path=None):
        """查找送货日期行"""
        valid_col_range = range(5, 12)
        invalid_labels = ["total", "total:", "sub-total", "sub-total:", "no. of cartons", "no. of cartons:"]
        found_blocks = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=max_rows)):
            cols = []
            for j, cell in enumerate(row):
                if j not in valid_col_range:
                    continue
                val = cell.value
                if cls.is_valid_date(val, next_week_start, next_week_end):
                    cols.append(j)
            if cols:
                found_blocks.append((i + 1, cols))
        for header_row, cols in found_blocks:
            for row_idx in range(header_row + 1, header_row + 100):
                label_val = ws.cell(row=row_idx, column=5).value
                label = str(label_val).lower().strip() if label_val else ""
                if any(invalid in label for invalid in invalid_labels):
                    continue
                for col_idx in cols:
                    qty_val = ws.cell(row=row_idx, column=col_idx+1).value
                    if isinstance(qty_val, (int, float)) and qty_val > 0:
                        return header_row, cols
        return None, []

    @classmethod
    def run_automation(cls, source_folder, supplier_folder, 
                      outlet_config=None, delivery_config=None,
                      log_callback=None, mapping_callback=None):
        """运行订单自动化（仅生成 supplier 文件）"""
        now = datetime.now()
        next_week_start = (now + timedelta(days=7 - now.weekday())).replace(hour=0, minute=0, second=0)
        next_week_end = next_week_start + timedelta(days=6)
        log_file = os.path.join(source_folder, f"order_log_{now.strftime('%Y%m%d_%H%M%S')}.txt")
        log_lines = [f"Scan Start Time: {now}\n", f"Target Week: {next_week_start.date()} to {next_week_end.date()}\n"]
        
        date_validator = DeliveryDateValidator(delivery_config) if delivery_config else None
        
        outlet_mapping = {}
        if outlet_config:
            try:
                outlet_mapping = {o['short_name']: o for o in outlet_config}
                if log_callback:
                    log_callback(f"✅ 已加载分店配置: {len(outlet_mapping)} 个分店")
            except Exception as e:
                if log_callback:
                    log_callback(f"⚠️ 分店配置加载失败: {str(e)}")
        
        # 新增：建立一個 OrderAutomation 實例用於 get_short_code
        oa_instance = cls(outlet_config) if outlet_config else cls()
        
        supplier_to_outlets = defaultdict(list)
        
        def log(message, include_timestamp=True):
            timestamp = datetime.now().strftime("%H:%M:%S") if include_timestamp else ""
            log_line = f"[{timestamp}] {message}" if include_timestamp else message
            log_lines.append(log_line + "\n")
            if log_callback:
                log_callback(log_line + "\n")
        
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
                log(f"工作表: {', '.join(wb.sheetnames)}")
                for sheetname in wb.sheetnames:
                    ws = wb[sheetname]
                    if ws.sheet_state != "visible":
                        log(f"  ⏩ 跳过隐藏工作表: {sheetname}")
                        continue
                    outlet_short = file.split('_')[0].strip() if '_' in file else file.split('.')[0].strip()
                    log(f"  Sheet: {sheetname}, Short Name (from filename): '{outlet_short}'")
                    has_order = False
                    # ========== 新 robust 檢測邏輯：允許 Mon-Sat/Mon-Fri/Mon-Sun 標題行 ==========
                    week_days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
                    for row in range(1, ws.max_row):
                        row_vals = [str(ws.cell(row=row, column=col).value).strip() if ws.cell(row=row, column=col).value else '' for col in range(6, 13)]
                        # 只要有 Mon 且有連續 3 個以上 weekday 就算標題行
                        weekday_count = sum(day in row_vals for day in week_days)
                        if 'Mon' in row_vals and weekday_count >= 3:
                            date_row = row + 1
                            date_cols = []
                            for idx, col in enumerate(range(6, 13)):
                                val = ws.cell(row=date_row, column=col).value
                                parsed = None
                                try:
                                    if val:
                                        sval = str(val)
                                        if re.match(r"^\d{1,2}-[A-Za-z]{3,}$", sval):
                                            year = next_week_start.year
                                            sval = f"{sval}-{year}"
                                        if isinstance(val, (int, float)):
                                            base_date = datetime(1899, 12, 30)
                                            parsed = base_date + timedelta(days=val)
                                        else:
                                            parsed = parse(sval, fuzzy=True, dayfirst=False)
                                except Exception:
                                    pass
                                if parsed and next_week_start.date() <= parsed.date() <= next_week_end.date():
                                    date_cols.append(col)
                            # 只保留關鍵 log
                            if date_cols:
                                for r in range(date_row+1, min(date_row+20, ws.max_row+1)):
                                    for c in date_cols:
                                        val = ws.cell(row=r, column=c).value
                                        if val is not None:
                                            sval = str(val).strip()
                                            try:
                                                fval = float(sval)
                                                if fval > 0:
                                                    has_order = True
                                                    break
                                            except:
                                                if sval:
                                                    has_order = True
                                                    break
                                    if has_order:
                                        break
                            if has_order:
                                break
                    if has_order:
                        log(f"    ✅ {outlet_short} has orders for {sheetname}")
                        supplier_to_outlets[sheetname].append((outlet_short, full_path, sheetname))
                        log(f"  {sheetname}: 有訂單: {outlet_short}", include_timestamp=False)
                    else:
                        log(f"    ⏩ No orders found for {outlet_short} in {sheetname}")
            except Exception as e:
                error_msg = f"❌ Error processing {file}: {str(e)}\n{traceback.format_exc()}"
                log(error_msg)
        
        log("\nSupplier-Outlet Mapping:")
        for supplier, outlets in supplier_to_outlets.items():
            outlet_list = [str(o[0]) if o[0] else "UNKNOWN" for o in outlets]
            log(f"  {supplier}: {', '.join(outlet_list)}")
        
        log("\nCreating supplier files...")
        supplier_files = []
        import xlwings as xw
        for sheetname, outlet_file_pairs in supplier_to_outlets.items():
            supplier_path = os.path.join(supplier_folder, f"{sheetname}_Week_{next_week_start.strftime('%V')}.xlsx")
            # 先建立空檔案
            if not os.path.exists(supplier_path):
                from openpyxl import Workbook
                wb = Workbook()
                wb.save(supplier_path)
            app = xw.App(visible=False)
            try:
                wb_dest = app.books.open(supplier_path)
                for outlet, src_file, original_sheet in outlet_file_pairs:
                    try:
                        wb_src = app.books.open(src_file)
                        sht = wb_src.sheets[original_sheet]
                        dest_sheet_name = outlet if outlet else original_sheet
                        # 刪除同名分頁
                        for s in wb_dest.sheets:
                            if s.name == dest_sheet_name:
                                s.delete()
                        sht.api.Copy(Before=wb_dest.sheets[0].api)
                        wb_dest.sheets[0].name = dest_sheet_name
                        wb_src.close()
                        log(f"  ✅ {dest_sheet_name} 已複製進 {os.path.basename(supplier_path)}（格式/公式完整保留）")
                    except Exception as e:
                        log(f"    ❌ Failed to copy {outlet} in {sheetname}: {str(e)}\n{traceback.format_exc()}")
                wb_dest.save()
                wb_dest.close()
                supplier_files.append(os.path.basename(supplier_path))
            finally:
                app.quit()
        result_text = "\n订单整合结果:\n"
        result_text += f"已处理供应商文件: {len(supplier_files)}\n\n"
        result_text += "供应商文件列表:\n" + "\n".join([f"- {file}" for file in supplier_files])
        log(result_text)
        try:
            with open(log_file, "w", encoding="utf-8") as logfile:
                logfile.writelines(log_lines)
            log(f"\nLog saved at: {log_file}")
            return True, f"订单整合完成！\n\n日志文件保存至:\n{log_file}\n\n{result_text}"
        except Exception as e:
            log(f"❌ Failed to write log file: {e}")
            return False, f"订单整合完成但日志保存失败:\n{str(e)}"

# ========== 增强版订单检查器 ==========
class EnhancedOrderChecker:
    """使用配置文件的订单检查器"""
    
    def __init__(self, config_manager=None):
        self.config_manager = config_manager
    
    @staticmethod
    def get_outlet_shortname(f5_value):
        """获取分店简称"""
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

    def run_checklist(self, folder, log_callback=None, as_table=False):
        date_validator = None
        if self.config_manager:
            delivery_schedules = self.config_manager.delivery_schedule
            if delivery_schedules:
                date_validator = DeliveryDateValidator()
                date_validator.schedule = defaultdict(dict)
                for row in delivery_schedules:
                    supplier = row['supplier']
                    outlet = row['outlet_code']
                    days = DeliveryDateValidator().parse_delivery_days(row['delivery_days'])
                    if outlet == "ALL":
                        date_validator.schedule[supplier]['*'] = days
                    else:
                        date_validator.schedule[supplier][outlet] = days
        supplier_keywords = {}
        if self.config_manager:
            for supplier in self.config_manager.suppliers:
                supplier_keywords[supplier["name"]] = [supplier["name"].lower()]
        must_have_outlets = {}
        if self.config_manager:
            for supplier in self.config_manager.suppliers:
                must_have_outlets[supplier["name"]] = self.config_manager.get_required_outlets(supplier["name"])
        try:
            files = [f for f in os.listdir(folder) if f.endswith(".xlsx") and not f.startswith("~$")]
            def file_key(f):
                import re
                base = f.split("_Week_")[0]
                return self._normalize(base)
            normalized_files = {file_key(f): f for f in files}
            output = []
            table = []
            for supplier, keywords in supplier_keywords.items():
                matches = self._find_supplier_file(normalized_files, keywords)
                if not matches:
                    if as_table:
                        table.append({
                            "supplier": supplier,
                            "outlet": "-",
                            "cover_status": "❌",
                            "date_status": "-",
                            "remark": "Supplier file not found"
                        })
                    else:
                        output.append(f"\n❌ {supplier} - Supplier file not found.")
                    continue
                for match in matches:
                    result, table_rows = self._process_supplier_file(
                        folder, match, must_have_outlets[supplier], 
                        date_validator, log_callback, as_table=True
                    )
                if as_table:
                    table.extend(table_rows)
                else:
                    output.extend(result)
            if as_table:
                return table
            return "\n".join(output)
        except Exception as e:
            if as_table:
                return [{"supplier": "-", "outlet": "-", "cover_status": "❌", "date_status": "-", "remark": f"Error: {str(e)}"}]
            return f"❌ Error running checklist: {str(e)}\n{traceback.format_exc()}"

    @staticmethod
    def _normalize(text):
        """标准化文本"""
        return re.sub(r'[\s\(\)]', '', text.lower())

    @classmethod
    def _find_supplier_file(cls, normalized_files, keywords):
        matches = []
        for k in keywords:
            nk = cls._normalize(k)
            for nf, of in normalized_files.items():
                if nf == nk:
                    matches.append(of)
        if matches:
            return matches
        # 若完全相等找不到，再用包含且唯一
        for k in keywords:
            nk = cls._normalize(k)
            partial = [of for nf, of in normalized_files.items() if nk in nf or nf in nk]
            if len(partial) == 1:
                return partial
        return None

    @classmethod
    def _process_supplier_file(cls, folder, filename, required_outlets, date_validator=None, log_callback=None, as_table=False):
        output = []
        table = []
        found = set()
        unidentified = []
        date_errors = []
        try:
            wb = load_workbook(os.path.join(folder, filename), data_only=True)
            for s in wb.sheetnames:
                try:
                    f5 = wb[s]["F5"].value
                    code = cls.get_outlet_shortname(f5)
                    if "[UNKNOWN]" in code or "[EMPTY]" in code:
                        unidentified.append((s, f5))
                        if as_table:
                            table.append({
                                "supplier": filename.split("_")[0],
                                "outlet": s,
                                "cover_status": "⚠️",
                                "date_status": "-",
                                "remark": f"Unidentified outlet: {f5}"
                            })
                    else:
                        found.add(code)
                        date_status = "-"
                        remark = ""
                        if date_validator:
                            delivery_date = wb[s]['F8'].value if 'F8' in wb[s] else None
                            if delivery_date:
                                is_valid = date_validator.validate_order(
                                    os.path.basename(filename).split('_')[0],
                                    code,
                                    delivery_date,
                                    log_callback
                                )
                                if not is_valid:
                                    date_status = "❌"
                                    remark = "Invalid delivery date"
                                    date_errors.append(code)
                                else:
                                    date_status = "✔️"
                            else:
                                date_status = "⚠️"
                                remark = "No delivery date"
                        if as_table:
                            table.append({
                                "supplier": filename.split("_")[0],
                                "outlet": code,
                                "cover_status": "✔️",
                                "date_status": date_status,
                                "remark": remark
                            })
                except:
                    unidentified.append((s, "[F5 error]"))
                    if as_table:
                        table.append({
                            "supplier": filename.split("_")[0],
                            "outlet": s,
                            "cover_status": "⚠️",
                            "date_status": "-",
                            "remark": "F5 error"
                        })
            required = set(required_outlets)
            for o in sorted(required - found):
                if as_table:
                    table.append({
                        "supplier": filename.split("_")[0],
                        "outlet": o,
                        "cover_status": "❌",
                        "date_status": "-",
                        "remark": "Missing outlet"
                    })
                else:
                    output.append(f"❌ {o}")
            if as_table:
                return output, table
            # 原本文字報表
            output.append(f"\n=== {filename} ===")
            output.append(f"📊 Required: {len(required)}, Found: {len(found)}, Missing: {len(required - found)}")
            for o in sorted(required & found):
                output.append(f"✔️ {o}")
            for o in sorted(required - found):
                output.append(f"❌ {o}")
            for s, v in unidentified:
                output.append(f"⚠️ {s} => {v}")
        except Exception as e:
            if as_table:
                table.append({"supplier": filename.split("_")[0], "outlet": "-", "cover_status": "❌", "date_status": "-", "remark": f"Error: {str(e)}"})
            else:
                output.append(f"❌ Error processing {filename}: {str(e)}")
        return output, table

# ========== 增强版订单自动化 ==========
class EnhancedOrderAutomation(OrderAutomation):
    """支持邮件发送的订单自动化"""
    
    def __init__(self, config_manager=None):
        super().__init__()
        self.config_manager = config_manager
        self.email_sender = EmailSender(config_manager)
    
    def run_automation(self, source_folder, supplier_folder, 
                      log_callback=None, mapping_callback=None, 
                      email_callback=None):
        """运行自动化流程（仅 supplier 整合）"""
        success, result = super().run_automation(
            source_folder, supplier_folder, 
            log_callback=log_callback, mapping_callback=mapping_callback
        )
        if not success:
            return success, result
        supplier_files = []
        for file in os.listdir(supplier_folder):
            if file.endswith(".xlsx") and "Week" in file:
                supplier_name = file.split("_")[0]
                supplier_files.append({
                    "path": os.path.join(supplier_folder, file),
                    "supplier": supplier_name
                })
        if self.config_manager and email_callback:
            email_callback(supplier_files)
        return success, result

# ========== 主应用程序 ==========
class SushiExpressApp(ctk.CTk):
    """主应用程序"""
    
    def __init__(self):
        super().__init__()
        self.title(f"Sushi Express Automation Tool v{VERSION}")
        self.geometry("1400x900")
        self.minsize(1200,800)
        self.configure(fg_color=DARK_BG)
        self.iconbitmap(resource_path("SELOGO22 - 01.ico"))
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # 初始化所有用到的 StringVar
        self.download_folder_var = ctk.StringVar()
        self.config_file_var = ctk.StringVar()
        self.checklist_folder_var = ctk.StringVar()
        self.master_config_var = ctk.StringVar()
        self.folder_vars = {}
        self.master_file_var = ctk.StringVar()
        self.output_folder_var = ctk.StringVar()
        self.email_supplier_folder_var = ctk.StringVar()
        self.email_master_config_var = ctk.StringVar()
        self.progress_popup = None
        self.mapping_popup = None
        self.outlet_config_var = ctk.StringVar()
        self.email_dialogs = []
        self.current_function = None
        self.nav_buttons = {}
        self.selected_outlook_account_idx = None
        self.checklist_search_var = ctk.StringVar()
        self._setup_ui()
        self.show_login()
    
    def _setup_ui(self):
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # 左側導航欄
        self.nav_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=DARK_PANEL,
            corner_radius=24,
            width=300
        )
        self.nav_frame.pack(side="left", fill="y", padx=(0, 10), pady=10)
        self.nav_frame.pack_propagate(False)

        # 加入 LOGO
        logo_img = load_image(LOGO_PATH, max_size=(220, 80))
        if logo_img:
            logo_label = ctk.CTkLabel(self.nav_frame, image=logo_img, text="")
            logo_label.image = logo_img  # 防止被垃圾回收
            logo_label.pack(pady=(18, 8))

        nav_title = ctk.CTkLabel(
            self.nav_frame,
            text="功能菜单\nFunction Menu",
            font=FONT_TITLE,
            text_color=ACCENT_BLUE,
            justify="center"
        )
        nav_title.pack(pady=20)
        ctk.CTkFrame(self.nav_frame, height=2, fg_color=ACCENT_BLUE).pack(fill="x", padx=20, pady=10)

        # 按鈕容器
        self.button_container = ctk.CTkFrame(self.nav_frame, fg_color="transparent")
        self.button_container.pack(fill="both", expand=True, padx=10, pady=10)

        # 右側內容區
        self.content_container = ctk.CTkFrame(
            self.main_container,
            fg_color=DARK_PANEL,
            corner_radius=24
        )
        self.content_container.pack(side="right", fill="both", expand=True, padx=10, pady=10)
        self.content_container.pack_propagate(False)

        # 右側內容區的標題區域
        self.content_header = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.content_header.pack(fill="x", padx=20, pady=20)

        self.function_title = ctk.CTkLabel(
            self.content_header, 
            text="", 
            font=FONT_TITLE,
            text_color=ACCENT_BLUE
        )
        self.function_title.pack(side="left")
        
        self.function_subtitle = ctk.CTkLabel(
            self.content_header, 
            text="", 
            font=FONT_SUB,
            text_color=TEXT_COLOR
        )
        self.function_subtitle.pack(side="left", padx=20)
        
        # 右侧内容区的主体
        self.content_body = ctk.CTkFrame(self.content_container, fg_color="transparent")
        self.content_body.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 添加返回主菜单按钮
        back_frame = ctk.CTkFrame(self.content_container, fg_color="transparent")
        back_frame.pack(fill="x", padx=20, pady=10)
        ctk.CTkButton(
            back_frame, 
            text="返回主菜单\nBack to Main Menu", 
            command=self.show_main_menu,
            fg_color=ACCENT_PURPLE,
            hover_color=BTN_HOVER
        ).pack(side="right")
    
    def _on_close(self):
        print("on_close called")
        self.destroy()
    
    def show_login(self):
        """显示登录界面"""
        for w in self.content_container.winfo_children():
            if w != self.content_header and w != self.content_body:
                w.destroy()
        
        # 清空内容区
        for w in self.content_body.winfo_children():
            w.destroy()
        
        # 创建登录界面
        login_frame = ctk.CTkFrame(
            self.content_body, 
            fg_color=DARK_PANEL, 
            corner_radius=24,
            width=450, 
            height=400
        )
        login_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        self.pwd_entry = ctk.CTkEntry(
            login_frame, 
            show="*", 
            font=FONT_MID,  # 改成較小字體
            width=300, 
            placeholder_text=t("password"), 
            fg_color=ENTRY_BG, 
            height=45
        )
        self.pwd_entry.pack(pady=(40,20))
        self.pwd_entry.bind("<Return>", lambda e: self._try_login())
        
        ctk.CTkButton(
            login_frame, 
            text=t("login_btn"), 
            command=self._try_login, 
            width=200,
            font=FONT_BIGBTN
        ).pack(pady=(0,20))
        
        ctk.CTkLabel(
            login_frame, 
            text=f"Version {VERSION} | {DEVELOPER}", 
            font=FONT_MID,
            text_color="#64748B"
        ).pack(side="bottom", pady=10)
        
        # 设置标题
        self.function_title.configure(text="系统登录\nSystem Login", font=FONT_TITLE)
        self.function_subtitle.configure(text="请输入密码進入系统\nPlease enter password to access the system", font=FONT_MID)
    
    def _try_login(self):
        """尝试登录"""
        if self.pwd_entry.get() == PASSWORD:
            self.show_main_menu()
        else:
            messagebox.showerror(t("error"), t("incorrect_pw"))
    
    def show_main_menu(self):
        # 清空内容区
        for w in self.content_body.winfo_children():
            w.destroy()

        # 设置标题
        self.function_title.configure(text=t("main_title"))
        self.function_subtitle.configure(text=t("select_function"))

        # 创建欢迎界面
        welcome_frame = ctk.CTkFrame(self.content_body, fg_color="transparent")
        welcome_frame.pack(fill="both", expand=True, padx=50, pady=50)

        welcome_text = (
            "欢迎使用Sushi Express 自动化工具\n"
            "Welcome to Sushi Express Automation Tool\n\n"
            "请从左侧菜单选择您需要的功能\n"
            "Please select a function from the left menu"
        )

        ctk.CTkLabel(
            welcome_frame,
            text=welcome_text,
            font=FONT_MID,
            text_color=TEXT_COLOR,
            justify="center"
        ).pack(expand=True)

        # **每次都重建按鈕**
        self._create_navigation_buttons()
    
    def _create_navigation_buttons(self):
        for widget in self.button_container.winfo_children():
            widget.destroy()
        inner_frame = ctk.CTkFrame(self.button_container, fg_color="transparent")
        inner_frame.pack(expand=True)
        functions = [
            ("download_title", self.show_download_ui, ACCENT_BLUE),
            ("automation_title", self.show_automation_ui, ACCENT_GREEN),
            ("checklist_title", self.show_checklist_ui, ACCENT_PURPLE),
            ("send_emails", self.show_email_sending_ui, ACCENT_RED),
            ("operation_supplies", self.show_operation_supplies_ui, "#F97316"),
            ("exit_system", self._on_close, "#64748B")
        ]
        for func_key, command, color in functions:
            btn = NavigationButton(
                inner_frame,
                text=t(func_key),
                command=command if func_key == "exit_system" else lambda c=command, k=func_key: self._select_function(c, k),
                fg_color=DARK_PANEL,
                anchor="center",
                border_width=2,
                border_color=ACCENT_BLUE,
                corner_radius=10,
                font=FONT_BIGBTN,
                text_color=get_contrast_color(DARK_PANEL),
                height=50
            )
            btn.pack(fill="x", pady=7, padx=10)
            self.nav_buttons[func_key] = btn
        # 新增：用戶指南按鈕
        help_btn = NavigationButton(
            inner_frame,
            text="用户指南\nUser Guide",
            command=self.show_user_guide,
            fg_color=ACCENT_PURPLE,
            anchor="center",
            border_width=2,
            border_color=ACCENT_PURPLE,
            corner_radius=10,
            font=FONT_BIGBTN,
            text_color=get_contrast_color(ACCENT_PURPLE),
            height=50
        )
        help_btn.pack(fill="x", pady=7, padx=10)
    
    def _select_function(self, command, func_key):
        """选择功能"""
        # 取消之前选中的按钮
        if self.current_function:
            self.nav_buttons[self.current_function].deselect()
        
        # 选中当前按钮
        self.nav_buttons[func_key].select()
        self.current_function = func_key
        
        # 执行功能
        command()
    
    def _show_function_ui(self, title_key, subtitle_key, content_callback):
        """显示功能界面"""
        # 清空内容区
        for w in self.content_body.winfo_children():
            w.destroy()
        
        # 设置标题
        self.function_title.configure(text=t(title_key))
        # 副標題統一放大
        if isinstance(subtitle_key, tuple):
            text = subtitle_key[0]
        else:
            text = t(subtitle_key)
        self.function_subtitle.configure(text=text, font=FONT_TITLE)
        
        # 创建内容框架
        content_frame = ctk.CTkFrame(self.content_body, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 构建功能UI
        content_callback(content_frame)
    
    # 下面是各个功能界面的修改，只需要将原来的show_xxx_ui方法改为使用_show_function_ui
    
    def show_download_ui(self):
        """显示下载界面"""
        def build(c):
            self.download_folder_var = ctk.StringVar()
            self.config_file_var = ctk.StringVar()
            
            # 创建表单框架
            form_frame = ctk.CTkFrame(c, fg_color="transparent")
            form_frame.pack(fill="both", expand=True, padx=50, pady=20)
            
            # 下载文件夹
            row1 = ctk.CTkFrame(form_frame)
            row1.pack(fill="x", pady=15)
            ctk.CTkLabel(row1, text="下载文件夹\nDownload Folder:", font=FONT_MID, anchor="w", justify="left").pack(side="left", padx=10)
            ctk.CTkEntry(row1, textvariable=self.download_folder_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text="浏览...\nBrowse...", font=FONT_MID, command=self._select_download_folder).pack(side="left", padx=5)
            
            # 分店配置文件
            row2 = ctk.CTkFrame(form_frame)
            row2.pack(fill="x", pady=15)
            ctk.CTkLabel(row2, text="分店配置文件（可选）\nOutlet Config (Optional):", font=FONT_MID, anchor="w", justify="left").pack(side="left", padx=10)
            ctk.CTkEntry(row2, textvariable=self.config_file_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text="浏览...\nBrowse...", font=FONT_MID, command=self._select_config_file).pack(side="left", padx=5)
            # 开始下载按钮
            btn_frame = ctk.CTkFrame(c, fg_color="transparent")
            btn_frame.pack(pady=30)
            GlowButton(
                btn_frame, 
                text="开始下载\nStart Download",
                command=self._run_download,
                width=300,
                height=60,
                glow_color=ACCENT_BLUE
            ).pack()
            # 确保"查看邮件内容"按钮始终显示，且更大
            def show_extracted_bodies():
                from tkinter import messagebox
                import os
                import tkinter as tk
                folder = self.download_folder_var.get()
                week_no = datetime.now().isocalendar()[1]
                save_path = os.path.join(folder, f"Week_{week_no}")
                log_file = os.path.join(save_path, "email_bodies_log.txt")
                bodies = getattr(OutlookDownloader, 'extracted_bodies', [])
                if not bodies and os.path.exists(log_file):
                    with open(log_file, "r", encoding="utf-8") as f:
                        content = f.read()
                    bodies = content.split("\n\n——— 邮件 ") if content else []
                    if bodies and not bodies[0].startswith("——— 邮件 "):
                        bodies[0] = "——— 邮件 1 ———\n" + bodies[0]
                    bodies = [b if b.startswith("——— 邮件 ") else "——— 邮件 " + b for b in bodies]
                if not bodies:
                    messagebox.showinfo("无内容", "请先下载邮件，再查看邮件内容！")
                    return
                # 搜索功能
                win = tk.Toplevel(self)
                win.title("邮件内容提取结果/Extracted Email Bodies")
                win.geometry("800x600")
                search_var = tk.StringVar()
                def filter_bodies(keyword):
                    keyword = keyword.lower().strip()
                    if not keyword:
                        return bodies
                    result = []
                    for b in bodies:
                        if keyword in b.lower():
                            result.append(b)
                    return result
                text_widget = tk.Text(win, wrap="word", font=("Microsoft JhengHei", 12))
                text_widget.pack(fill="both", expand=True, padx=10, pady=(40,10))
                def update_display():
                    filtered = filter_bodies(search_var.get())
                    pretty = []
                    for idx, b in enumerate(filtered, 1):
                        lines = b.split("\n")
                        sender = lines[0] if lines else ""
                        subject = lines[1] if len(lines) > 1 else ""
                        items = lines[2:] if len(lines) > 2 else []
                        if not items or all(not i.strip() for i in items):
                            items = ["无数字开头内容"]
                        pretty.append(f"——— 邮件 {idx} ———\n[发件人] {sender}\n[主题] {subject}\n[内容]")
                        for i, item in enumerate(items, 1):
                            pretty.append(f"  {i}. {item.strip()}")
                    pretty_text = "\n\n".join(pretty)
                    text_widget.delete("1.0", "end")
                    text_widget.insert("1.0", pretty_text)
                def on_search(*args):
                    update_display()
                search_var.trace_add("write", on_search)
                search_entry = tk.Entry(win, textvariable=search_var, font=("Microsoft JhengHei", 13))
                search_entry.place(x=10, y=5, width=400, height=30)
                search_entry.insert(0, "输入关键字搜索/Type to search...")
                def on_focus_in(event):
                    if search_entry.get() == "输入关键字搜索/Type to search...":
                        search_entry.delete(0, "end")
                search_entry.bind("<FocusIn>", on_focus_in)
                # 支持复制
                def copy_all():
                    import pyperclip
                    pretty = text_widget.get("1.0", "end-1c")
                    pyperclip.copy(pretty)
                    messagebox.showinfo("复制成功", "已复制全部邮件内容！")
                btn = tk.Button(win, text="复制全部内容\nCopy All", command=copy_all, bg="#22c55e", fg="#fff", font=("Microsoft JhengHei", 14, "bold"))
                btn.place(x=420, y=5, width=180, height=30)
                update_display()
            GlowButton(
                btn_frame,
                text="查看邮件内容\nCheck Email Bodies",
                command=show_extracted_bodies,
                width=300,
                height=48,
                glow_color=ACCENT_GREEN
            ).pack(pady=10)
        
        self._show_function_ui(
            "download_title", 
            "download_desc", 
            build
        )
    
    def show_checklist_ui(self):
        """显示检查表界面"""
        def build(c):
            self.checklist_folder_var = ctk.StringVar()
            self.master_config_var = ctk.StringVar()
            # 创建表单框架
            form_frame = ctk.CTkFrame(c, fg_color="transparent")
            form_frame.pack(fill="both", expand=True, padx=50, pady=5)
            # 订单文件夹
            row1 = ctk.CTkFrame(form_frame)
            row1.pack(fill="x", pady=5)
            ctk.CTkLabel(row1, text="订单文件夹/Order Folder:", font=FONT_MID).pack(side="left", padx=10)
            ctk.CTkEntry(row1, textvariable=self.checklist_folder_var, font=FONT_MID, state="readonly", width=250).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=str(t("browse")), font=FONT_MID, command=self._select_checklist_folder).pack(side="left", padx=5)
            # 统一配置文件
            row2 = ctk.CTkFrame(form_frame)
            row2.pack(fill="x", pady=5)
            ctk.CTkLabel(row2, text="统一配置文件/Master Config (Excel):", font=FONT_MID).pack(side="left", padx=10)
            ctk.CTkEntry(row2, textvariable=self.master_config_var, font=FONT_MID, state="readonly", width=250).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=str(t("browse")), font=FONT_MID, command=lambda: self._select_config_file(self.master_config_var, [("Excel files", "*.xlsx")])).pack(side="left", padx=5)
            # 搜索框
            search_frame = ctk.CTkFrame(c, fg_color="transparent")
            search_frame.pack(fill="x", pady=5)
            ctk.CTkLabel(search_frame, text="搜索/Filter:", font=FONT_MID).pack(side="left", padx=10)
            if not hasattr(self, 'checklist_search_var'):
                self.checklist_search_var = ctk.StringVar()
            search_entry = ctk.CTkEntry(search_frame, textvariable=self.checklist_search_var, font=FONT_MID, width=300)
            search_entry.pack(side="left", padx=5)
            search_entry.bind("<KeyRelease>", lambda e: self._filter_checklist_table())
            # 表格區
            table_frame = ctk.CTkFrame(c, fg_color="transparent")
            table_frame.pack(fill="both", expand=True, padx=10, pady=10)
            import tkinter.ttk as ttk
            style = ttk.Style()
            style.theme_use('default')
            style.configure("Custom.Treeview", background="#1e293b", fieldbackground="#1e293b", foreground="#e2e8f0", rowheight=28, font=("Microsoft JhengHei", 12))
            style.configure("Custom.Treeview.Heading", background="#334155", foreground="#60a5fa", font=("Microsoft JhengHei", 13, "bold"))
            style.map("Custom.Treeview", background=[('selected', '#334155')])
            if not hasattr(self, 'checklist_table'):
                self.checklist_table = None
            self.checklist_table = ttk.Treeview(
                table_frame, 
                columns=("supplier", "outlet", "cover_status", "date_status", "remark"), 
                show="headings", 
                height=8,
                style="Custom.Treeview"
            )
            col_labels = [
                ("supplier", "供應商/Supplier"),
                ("outlet", "分店/Outlet"),
                ("cover_status", "覆蓋狀態/Cover"),
                ("date_status", "日期狀態/Date"),
                ("remark", "備註/Remark")
            ]
            for col, label in col_labels:
                self.checklist_table.heading(col, text=label)
                self.checklist_table.column(col, width=140 if col!="remark" else 300, anchor="center")
            self.checklist_table.pack(fill="both", expand=True)
            self.checklist_table.bind("<Double-1>", self._on_checklist_row_double_click)
            # 複製/匯出按鈕
            btn_frame = ctk.CTkFrame(c, fg_color="transparent")
            btn_frame.pack(pady=5)
            ctk.CTkButton(btn_frame, text="複製報表/Copy", command=self._copy_checklist_table).pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="匯出Excel/Export Excel", command=self._export_checklist_table).pack(side="left", padx=10)
            # 新增查看必要門市按鈕
            ctk.CTkButton(btn_frame, text="查看必要門市/View Required Outlets", command=self._show_required_outlets_window).pack(side="left", padx=10)
            # 新增 Cross Check 按鈕
            ctk.CTkButton(btn_frame, text="交叉檢查\nCross Check", command=self._run_cross_check_email_log).pack(side="left", padx=10)
            # 執行檢查按鈕置中
            run_btn_frame = ctk.CTkFrame(c, fg_color="transparent")
            run_btn_frame.pack(fill="x", pady=15)
            GlowButton(
                run_btn_frame, 
                text="執行檢查/Run Check",
                command=self._run_enhanced_checklist,
                width=220,
                height=48,
                glow_color=ACCENT_PURPLE
            ).pack(anchor="center")
            # 初始化表格數據
            self._checklist_table_data = []
        self._show_function_ui(
            "checklist_title", 
            "checklist_desc", 
            build
        )
    def _run_enhanced_checklist(self):
        folder = self.checklist_folder_var.get()
        master_config = self.master_config_var.get()
        if not folder or not master_config:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        config_mgr = UnifiedConfigManager(master_config)
        checker = EnhancedOrderChecker(config_mgr)
        table = checker.run_checklist(folder, as_table=True)
        self._checklist_table_data = table
        self._refresh_checklist_table()
        # ====== 新增 cross check with email log ======
        import os
        email_log_path = os.path.join(folder, "email_bodies_log.txt")
        if os.path.exists(email_log_path):
            with open(email_log_path, "r", encoding="utf-8") as f:
                content = f.read()
            import re
            # 假設格式：——— 郵件 N ———\n[發件人] ...\n[主題] ...\n[內容]\n 1. Aries\n 2. Changcheng ...
            cross_check_result = []
            for block in content.split("——— 郵件"):
                lines = block.strip().splitlines()
                if not lines or len(lines) < 3:
                    continue
                outlet = lines[1].replace("[發件人]", "").strip() if lines[1].startswith("[發件人]") else ""
                claimed = []
                for line in lines[3:]:
                    m = re.match(r"\d+\.\s*(.+)", line)
                    if m:
                        claimed.append(m.group(1).strip())
                # cross check: 檢查 claimed 是否有出現在 checklist table
                found = set()
                for row in table:
                    if outlet and outlet in row["outlet"] and row["supplier"] in claimed and row["cover_status"] == "✔️":
                        found.add(row["supplier"])
                missed = [s for s in claimed if s not in found]
                if missed:
                    cross_check_result.append(f"[警告] {outlet} 聲稱下單但未整合到: {', '.join(missed)}")
            if cross_check_result:
                messagebox.showwarning("Cross Check Result", "\n".join(cross_check_result))
            else:
                messagebox.showinfo("Cross Check Result", "所有門市聲稱下單的供應商都已整合！")
        else:
            messagebox.showinfo("Cross Check Result", "找不到 email_bodies_log.txt！")
    def _refresh_checklist_table(self):
        # 清空表格
        for row in self.checklist_table.get_children():
            self.checklist_table.delete(row)
        # 插入數據
        for row in self._checklist_table_data:
            values = (row["supplier"], row["outlet"], row["cover_status"], row["date_status"], row["remark"])
            self.checklist_table.insert("", "end", values=values)
    def _filter_checklist_table(self):
        keyword = self.checklist_search_var.get().lower()
        def normalize(text):
            import re
            return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()
        norm_keyword = normalize(keyword)
        filtered = [row for row in self._checklist_table_data if norm_keyword in normalize(row["supplier"]) or norm_keyword in normalize(row["outlet"]) or norm_keyword in row["cover_status"] or norm_keyword in row["date_status"] or norm_keyword in row["remark"].lower()]
        for row in self.checklist_table.get_children():
            self.checklist_table.delete(row)
        for row in filtered:
            values = (row["supplier"], row["outlet"], row["cover_status"], row["date_status"], row["remark"])
            self.checklist_table.insert("", "end", values=values)
    def _copy_checklist_table(self):
        import pyperclip
        rows = ["\t".join(["供应商", "分店", "覆蓋狀態", "日期狀態", "備註"])]
        for row in self._checklist_table_data:
            rows.append("\t".join([str(row["supplier"]), str(row["outlet"]), str(row["cover_status"]), str(row["date_status"]), str(row["remark"])]))
        pyperclip.copy("\n".join(rows))
        messagebox.showinfo("複製成功", "報表已複製到剪貼簿！")
    def _export_checklist_table(self):
        import pandas as pd
        from tkinter import filedialog
        df = pd.DataFrame(self._checklist_table_data)
        file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file:
            df.to_excel(file, index=False)
            messagebox.showinfo("匯出成功", f"已匯出到 {file}")
    def _on_checklist_row_double_click(self, event):
        item = self.checklist_table.selection()
        if not item:
            return
        values = self.checklist_table.item(item[0], "values")
        detail = f"供应商: {values[0]}\n分店: {values[1]}\n覆蓋狀態: {values[2]}\n日期狀態: {values[3]}\n備註: {values[4]}"
        ScrollableMessageBox(self, "详细信息", detail)
    
    def show_automation_ui(self):
        """显示自动化界面"""
        def build(c):
            self.folder_vars = {}
            self.master_config_var = ctk.StringVar()
            # 创建表单框架
            form_frame = ctk.CTkFrame(c, fg_color="transparent")
            form_frame.pack(side="top", fill="x", expand=False, padx=50, pady=20)
            # 两个文件夹选择（移除分店文件夹）
            folders = [
                ("source_folder", "来源文件夹\nSource Folder (Weekly Orders)"),
                ("supplier_folder", "供应商文件夹\nSupplier Folder")
            ]
            for key, label in folders:
                row = ctk.CTkFrame(form_frame)
                row.pack(fill="x", pady=15)
                var = ctk.StringVar()
                self.folder_vars[key] = var
                ctk.CTkLabel(row, text=label, font=FONT_MID, anchor="w", justify="left").pack(side="left", padx=10)
                ctk.CTkEntry(row, textvariable=var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
                ctk.CTkButton(row, text="浏览...\nBrowse...", font=FONT_MID, command=lambda k=key: self._select_folder(k)).pack(side="left", padx=5)
            # 统一配置文件
            row3 = ctk.CTkFrame(form_frame)
            row3.pack(fill="x", pady=15)
            ctk.CTkLabel(row3, text="统一配置文件\nMaster Config (Excel):", font=FONT_MID, anchor="w", justify="left").pack(side="left", padx=10)
            ctk.CTkEntry(row3, textvariable=self.master_config_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row3, text="浏览...\nBrowse...", font=FONT_MID, command=lambda: self._select_config_file(self.master_config_var, [("Excel files", "*.xlsx")])).pack(side="left", padx=5)
            # 開始自動化按鈕（往上移，緊貼表單）
            btn_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
            btn_frame.pack(pady=(20, 0))
            GlowButton(
                btn_frame, 
                text="开始整合档案\nStart Automation",
                command=self._run_enhanced_automation,
                width=300,
                height=60,
                glow_color=ACCENT_GREEN
            ).pack()
            # ========== 補單區塊 ==========
            append_frame = ctk.CTkFrame(c, fg_color=DARK_PANEL, corner_radius=16)
            append_frame.pack(side="top", fill="x", expand=False, padx=50, pady=(10, 0))
            ctk.CTkLabel(append_frame, text="補加門市訂單\nAppend Outlet Order", font=FONT_BIGBTN, text_color=ACCENT_BLUE, anchor="w", justify="left").pack(anchor="w", padx=10, pady=(10,5))
            # 選擇要補的門市訂單檔案（可多選）
            self.append_outlet_files_var = ctk.StringVar()
            def select_append_files():
                from tkinter import filedialog
                files = filedialog.askopenfilenames(title="選擇要補的門市訂單/Select Outlet Order Files", filetypes=[("Excel files", "*.xlsx")])
                if files:
                    self.append_outlet_files_var.set(";".join(files))
            rowa = ctk.CTkFrame(append_frame)
            rowa.pack(fill="x", pady=5)
            ctk.CTkLabel(rowa, text="選擇要補的門市訂單檔案\nSelect Outlet Order Files:", font=FONT_MID, anchor="w", justify="left").pack(side="left", padx=10)
            ctk.CTkEntry(rowa, textvariable=self.append_outlet_files_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(rowa, text="浏览...\nBrowse...", font=FONT_MID, command=select_append_files).pack(side="left", padx=5)
            # 選擇已整合的Supplier檔案資料夾
            self.append_supplier_folder_var = ctk.StringVar()
            rowb = ctk.CTkFrame(append_frame)
            rowb.pack(fill="x", pady=5)
            ctk.CTkLabel(rowb, text="選擇已整合的Supplier檔案資料夾\nSelect Supplier Folder:", font=FONT_MID, anchor="w", justify="left").pack(side="left", padx=10)
            ctk.CTkEntry(rowb, textvariable=self.append_supplier_folder_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(rowb, text="浏览...\nBrowse...", font=FONT_MID, command=lambda: self._select_folder_var(self.append_supplier_folder_var)).pack(side="left", padx=5)
            # 補單按鈕（加大尺寸，往上緊貼欄位）
            def run_append_order():
                import os
                import copy
                import re
                from tkinter import messagebox
                from openpyxl import load_workbook
                from dateutil.parser import parse
                import xlwings as xw
                outlet_files = self.append_outlet_files_var.get().split(';')
                supplier_folder = self.append_supplier_folder_var.get()
                if not outlet_files or not supplier_folder:
                    messagebox.showwarning("警告", "請選擇要補的門市訂單檔案和 Supplier 資料夾！")
                    return
                summary = []
                def normalize(text):
                    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()
                def copy_sheet_xlwings(src_file, sheet_name, dest_file, dest_sheet_name=None):
                    app = xw.App(visible=False)
                    try:
                        wb_src = app.books.open(src_file)
                        wb_dest = app.books.open(dest_file)
                        sht = wb_src.sheets[sheet_name]
                        if dest_sheet_name is None:
                            dest_sheet_name = sheet_name
                        for s in wb_dest.sheets:
                            if s.name == dest_sheet_name:
                                s.delete()
                        sht.api.Copy(Before=wb_dest.sheets[0].api)
                        wb_dest.sheets[0].name = dest_sheet_name
                        wb_dest.save()
                        wb_src.close()
                        wb_dest.close()
                    finally:
                        app.quit()
                for outlet_file in outlet_files:
                    if not outlet_file or not os.path.exists(outlet_file):
                        continue
                    try:
                        base = os.path.basename(outlet_file)
                        m = re.match(r'([A-Za-z0-9]+)', base)
                        if not m:
                            summary.append(f"❌ {base} 檔名無法解析門市代碼")
                            continue
                        outlet_code = m.group(1)
                        week_num = None
                        wb_outlet = load_workbook(outlet_file, data_only=True)
                        for sheetname in wb_outlet.sheetnames:
                            ws_outlet = wb_outlet[sheetname]
                            if ws_outlet.sheet_state != "visible":
                                summary.append(f"⏩ 跳過隱藏工作表: {sheetname}")
                                continue
                            supplier_name = sheetname.strip()
                            now = datetime.now()
                            next_week_start = (now + timedelta(days=7 - now.weekday())).replace(hour=0, minute=0, second=0)
                            next_week_end = next_week_start + timedelta(days=6)
                            has_order = False
                            row = 1
                            week_days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
                            found_date_info = False
                            while row <= ws_outlet.max_row:
                                f_val = ws_outlet.cell(row=row, column=6).value
                                if f_val and str(f_val).strip() == 'Mon':
                                    header_cells = [str(ws_outlet.cell(row=row, column=col).value).strip() if ws_outlet.cell(row=row, column=col).value else '' for col in range(6, 12+1)]
                                    if header_cells[:6] == week_days:
                                        date_row = row + 1
                                        date_cells = []
                                        for col in range(6, 13):
                                            val = ws_outlet.cell(row=date_row, column=col).value
                                            if val is None:
                                                for up in range(date_row-1, 0, -1):
                                                    up_val = ws_outlet.cell(row=up, column=col).value
                                                    if up_val is not None:
                                                        val = up_val
                                                        break
                                            parsed = None
                                            try:
                                                if val:
                                                    sval = str(val)
                                                    if re.match(r"^\d{1,2}-[A-Za-z]{3,}$", sval):
                                                        year = next_week_start.year
                                                        sval = f"{sval}-{year}"
                                                    if isinstance(val, (int, float)):
                                                        base_date = datetime(1899, 12, 30)
                                                        parsed = base_date + timedelta(days=val)
                                                    else:
                                                        parsed = parse(sval, fuzzy=True, dayfirst=False)
                                            except Exception as e:
                                                if val and isinstance(val, str) and re.search(r"jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|mon|tue|wed|thu|fri|sat|sun", val, re.I):
                                                    print(f"[DEBUG] 日期解析失敗 row={date_row} col={col} val={val} err={e}")
                                            date_cells.append(parsed)
                                        # 顯示找到日期行與欄
                                        if not found_date_info:
                                            col_letters = [chr(64+col) for col in range(6, 13)]
                                            summary.append(f"Found delivery dates at row {date_row}, columns {col_letters}")
                                            found_date_info = True
                                        if any(d and next_week_start.date() <= d.date() <= next_week_end.date() for d in date_cells):
                                            qty_row = date_row + 1
                                            for col_idx, d in enumerate(date_cells):
                                                if d and next_week_start.date() <= d.date() <= next_week_end.date():
                                                    qty_val = ws_outlet.cell(row=qty_row, column=6+col_idx).value
                                                    if isinstance(qty_val, (int, float)) and qty_val > 0:
                                                        has_order = True
                                                        break
                                            if has_order:
                                                break
                                        row = row + 3
                                        continue
                                row += 1
                            if has_order:
                                summary.append(f"✅ {outlet_code} has orders for {sheetname}.")
                            else:
                                summary.append(f"⏩ {base} 的 {sheetname} 下周無訂單，未補單")
                                continue
                            # 在 supplier folder 找對應週別的 supplier 檔案
                            supplier_files = [f for f in os.listdir(supplier_folder) if f.endswith('.xlsx') and not f.startswith('~$')]
                            norm_supplier = normalize(supplier_name)
                            matched_file = None
                            for f in supplier_files:
                                m2 = re.match(r'(.+)_Week[_-]([0-9]+)\.xlsx', f, re.IGNORECASE)
                                if not m2:
                                    continue
                                fbase = m2.group(1)
                                if normalize(fbase) == norm_supplier:
                                    matched_file = f
                                    break
                            if not matched_file:
                                summary.append(f"❌ {base} 的 {sheetname} 找不到對應的 Supplier 檔案（{supplier_name}）")
                                continue
                            supplier_path = os.path.join(supplier_folder, matched_file)
                            # 用 xlwings 複製分頁，保留格式/公式/合併儲存格
                            try:
                                copy_sheet_xlwings(outlet_file, sheetname, supplier_path, dest_sheet_name=outlet_code)
                                summary.append(f"✅ {base} 的 {sheetname} 已補進 {matched_file}（格式/公式完整保留）")
                            except Exception as e:
                                summary.append(f"❌ {base} 的 {sheetname} 複製失敗: {str(e)}")
                    except Exception as e:
                        summary.append(f"❌ {base} 發生錯誤: {str(e)}")
                # 用自訂滾動視窗顯示結果
                ScrollableMessageBox(self, "補單結果", "\n".join(summary) if summary else "沒有任何檔案被處理")
            GlowButton(
                append_frame,
                text="開始補單\nAppend",
                command=run_append_order,
                width=300,
                height=60,
                glow_color=ACCENT_GREEN
            ).pack(pady=(10, 0))
        
        self._show_function_ui(
            "automation_title", 
            "automation_desc", 
            build
        )
    
    def show_email_sending_ui(self):
        """显示邮件发送界面"""
        def build(c):
            import os
            self.email_supplier_folder_var = ctk.StringVar()
            self.email_master_config_var = ctk.StringVar()
            # 创建表单框架
            form_frame = ctk.CTkFrame(c, fg_color="transparent")
            form_frame.pack(fill="both", expand=True, padx=50, pady=20)
            # 供应商文件夹
            row1 = ctk.CTkFrame(form_frame)
            row1.pack(fill="x", pady=15)
            ctk.CTkLabel(row1, text="供应商文件夹/Supplier Folder:", font=FONT_MID).pack(side="left", padx=10)
            ctk.CTkEntry(row1, textvariable=self.email_supplier_folder_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=t("browse"), font=FONT_MID, command=lambda: self._select_folder_var(self.email_supplier_folder_var)).pack(side="left", padx=5)
            # 统一配置文件
            row2 = ctk.CTkFrame(form_frame)
            row2.pack(fill="x", pady=15)
            ctk.CTkLabel(row2, text="统一配置文件/Master Config (Excel):", font=FONT_MID).pack(side="left", padx=10)
            ctk.CTkEntry(row2, textvariable=self.email_master_config_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=t("browse"), font=FONT_MID, command=lambda: self._select_config_file(self.email_master_config_var, [("Excel files", "*.xlsx")])).pack(side="left", padx=5)
            # 新增：顯示目前選擇的Outlook帳號
            account_frame = ctk.CTkFrame(form_frame, fg_color=ACCENT_BLUE, corner_radius=8)
            account_frame.pack(fill="x", pady=10, padx=5)
            self.selected_account_label_var = ctk.StringVar()
            def update_account_label():
                if hasattr(self, 'selected_outlook_account_idx') and self.selected_outlook_account_idx is not None:
                    try:
                        import win32com.client
                        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                        accounts = [outlook.Folders.Item(i + 1) for i in range(outlook.Folders.Count)]
                        if 0 <= self.selected_outlook_account_idx < len(accounts):
                            name = accounts[self.selected_outlook_account_idx].Name
                            self.selected_account_label_var.set(f"当前使用邮箱/Current Email: {name}")
                            return
                    except Exception:
                        pass
                self.selected_account_label_var.set("当前使用邮箱/Current Email: 未选定/Not selected")
            update_account_label()
            ctk.CTkLabel(account_frame, textvariable=self.selected_account_label_var, font=FONT_MID, text_color="#fff").pack(anchor="w", padx=10, pady=5)
            self.update_account_label = update_account_label  # 供外部調用
            # 郵件正文body輸入框
            ctk.CTkLabel(form_frame, text="郵件正文/Email Body:", font=FONT_MID).pack(anchor="w", padx=10, pady=(20,0))
            self.email_body_textbox = ctk.CTkTextbox(form_frame, height=120, font=FONT_MID, wrap="word")
            self.email_body_textbox.pack(fill="x", padx=10, pady=(0,10))
            # 載入本地body預設
            default_body = "Hi Team,<br><br>Week {week_no} order as attached.<br><br>"
            body_path = "email_body.txt"
            if os.path.exists(body_path):
                try:
                    with open(body_path, "r", encoding="utf-8") as f:
                        saved_body = f.read()
                    self.email_body_textbox.delete("1.0", "end")
                    self.email_body_textbox.insert("1.0", saved_body)
                except Exception:
                    self.email_body_textbox.insert("1.0", default_body)
            else:
                self.email_body_textbox.insert("1.0", default_body)
            # 保存按鈕
            def save_body():
                body = self.email_body_textbox.get("1.0", "end-1c")
                try:
                    with open("email_body.txt", "w", encoding="utf-8") as f:
                        f.write(body)
                    messagebox.showinfo("保存成功", "郵件正文已保存！")
                except Exception as e:
                    messagebox.showerror("保存失敗", f"保存失敗: {e}")
            ctk.CTkButton(form_frame, text="保存郵件正文/Save Email Body", command=save_body, fg_color=ACCENT_GREEN, hover_color=BTN_HOVER).pack(anchor="e", padx=10, pady=(0,10))
            # 发送邮件按钮
            btn_frame = ctk.CTkFrame(c, fg_color="transparent")
            btn_frame.pack(pady=30)
            GlowButton(
                btn_frame, 
                text="发送邮件给供应商/Send Emails to Suppliers",
                command=self._send_supplier_emails,
                width=300,
                height=60,
                glow_color=ACCENT_RED
            ).pack()
        self._show_function_ui(
            "send_emails",
            ("选择供应商文件夹并发送邮件\nSelect supplier folder and send emails", FONT_TITLE),
            build
        )
    def _send_supplier_emails(self):
        import os
        import re
        folder = self.email_supplier_folder_var.get()
        master_config = self.email_master_config_var.get()
        if not folder or not master_config:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        # 選擇Outlook帳號（只選一次）
        if self.selected_outlook_account_idx is None:
            try:
                import win32com.client
                outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                accounts = [outlook.Folders.Item(i + 1) for i in range(outlook.Folders.Count)]
                account_names = [acct.Name for acct in accounts]
                from tkinter import simpledialog
                idx = simpledialog.askinteger(
                    "選擇Outlook帳號/Select Outlook Account",
                    "\n".join([f"[{i}] {name}" for i, name in enumerate(account_names)]) + "\n請輸入序號/Please enter index:",
                    minvalue=0, maxvalue=len(account_names)-1,
                    parent=self
                )
                if idx is None:
                    messagebox.showwarning("取消/Cancelled", "未選擇帳號，已取消發送/No account selected, sending cancelled.")
                    return
                self.selected_outlook_account_idx = idx
            except Exception as e:
                messagebox.showerror("錯誤/Error", f"獲取Outlook帳號失敗/Failed to get Outlook accounts: {e}")
            return
        # 取得body內容，並保存到本地
        body = self.email_body_textbox.get("1.0", "end-1c")
        try:
            with open("email_body.txt", "w", encoding="utf-8") as f:
                f.write(body)
        except Exception:
            pass
        files = [f for f in os.listdir(folder) if f.endswith(".xlsx") and not f.startswith("~$")]
        config_mgr = UnifiedConfigManager(master_config)
        email_sender = EmailSender(config_mgr)
        supplier_names = set()
        wb = load_workbook(master_config, data_only=True)
        if "Suppliers" in wb.sheetnames:
            ws = wb["Suppliers"]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0]:
                    supplier_names.add(str(row[0]).strip())
        def normalize(name):
            return re.sub(r'[^a-zA-Z0-9]', '', name).lower()
        email_list = []
        matched_files = set()
        for file in files:
            file_base = file.split("_Week_")[0]
            nf = normalize(file_base)
            matched_supplier = None
            for sname in supplier_names:
                ns = normalize(sname)
                if ns in nf or nf in ns:
                    matched_supplier = sname
                    break
            if not matched_supplier:
                continue
            matched_files.add(file)
            supplier_file = os.path.join(folder, file)
            to_emails, cc_emails = email_sender.get_to_cc_emails(matched_supplier, master_config)
            now = datetime.now()
            week_no = now.isocalendar()[1]
            # 用body框內容，支援{week_no}自動替換
            body_filled = body.replace("{week_no}", str(week_no))
            mail = email_sender.send_email(
                to_emails, cc_emails, matched_supplier, body_filled,
                attachment_path=supplier_file,
                account_idx=self.selected_outlook_account_idx
            )
            if isinstance(mail, tuple):
                messagebox.showerror("Error", mail[1])
                continue
            if not mail:
                messagebox.showerror("Error", f"无法创建邮件: {matched_supplier}")
                continue
            email_list.append((mail, matched_supplier, supplier_file))
        # 額外提示未配對到的檔案
        unmatched_files = [f for f in files if f not in matched_files]
        if unmatched_files:
            messagebox.showwarning("未配對檔案", f"以下檔案未配對到任何Supplier，未發送郵件：\n" + "\n".join(unmatched_files))
        for mail, matched_supplier, supplier_file in email_list:
            def on_confirm(success, msg, ms=matched_supplier):
                if success:
                    messagebox.showinfo("Success", f"{ms} 郵件已發送！")
                else:
                    messagebox.showinfo("已取消/Cancelled", f"{ms} 郵件未發送。")
            EmailConfirmationDialog(self, mail, matched_supplier, "-", supplier_file, on_confirm)
    
    def show_operation_supplies_ui(self):
        """显示运营用品界面"""
        def build(c):
            self.master_file_var = ctk.StringVar()
            self.output_folder_var = ctk.StringVar()
            
            # 创建表单框架
            form_frame = ctk.CTkFrame(c, fg_color="transparent")
            form_frame.pack(fill="both", expand=True, padx=50, pady=20)
            
            # 主文件选择
            row1 = ctk.CTkFrame(form_frame)
            row1.pack(fill="x", pady=15)
            ctk.CTkLabel(row1, text="选择主文件/Select Master File:", font=FONT_MID).pack(side="left", padx=10)
            ctk.CTkEntry(row1, textvariable=self.master_file_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=t("browse"), font=FONT_MID, command=self._select_master_file).pack(side="left", padx=5)
            
            # 输出文件夹选择
            row2 = ctk.CTkFrame(form_frame)
            row2.pack(fill="x", pady=15)
            ctk.CTkLabel(row2, text="选择输出文件夹/Select Output Folder:", font=FONT_MID).pack(side="left", padx=10)
            ctk.CTkEntry(row2, textvariable=self.output_folder_var, font=FONT_MID, state="readonly", width=400).pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=t("browse"), font=FONT_MID, command=self._select_output_folder).pack(side="left", padx=5)
            
            # 开始处理按钮
            btn_frame = ctk.CTkFrame(c, fg_color="transparent")
            btn_frame.pack(pady=30)
            GlowButton(
                btn_frame, 
                text="开始处理/Start Processing",
                command=self._run_operation_supplies,
                width=300,
                height=60,
                glow_color="#F97316"
            ).pack()
        
        self._show_function_ui(
            "operation_supplies",
            ("处理运营用品月订单\nProcess monthly operation supplies orders", FONT_TITLE),
            build
        )

    def _select_master_file(self):
        """选择主文件"""
        f = filedialog.askopenfilename(title=t("select_folder"), filetypes=[("Excel","*.xlsx;*.xls")])
        if f: self.master_file_var.set(f)

    def _select_output_folder(self):
        """选择输出文件夹"""
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.output_folder_var.set(f)

    def _run_operation_supplies(self):
        """运行运营用品处理"""
        mf, of = self.master_file_var.get(), self.output_folder_var.get()
        if not mf or not of: messagebox.showwarning(t("warning"), t("folder_warning")); return
        p = ProgressPopup(self, t("processing"))
        self.progress_popup = p
        def logcb(m): p.log(m)
        threading.Thread(target=lambda: self._thread_task(lambda: OperationSuppliesOrder.process_order(mf, of, logcb))).start()

    def _thread_task(self, fn, show_message=True):
        """线程任务"""
        try:
            result = fn()
            if show_message:
                self.after(0, lambda: messagebox.showinfo(t("success"), result))
        except Exception as e:
            self.after(0, lambda: messagebox.showerror(t("error"), f"操作失败: {str(e)}"))

    def _select_download_folder(self):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f:
            self.download_folder_var.set(f)

    def _select_checklist_folder(self):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f:
            self.checklist_folder_var.set(f)

    def _run_download(self):
        folder = self.download_folder_var.get()
        config = self.config_file_var.get()
        if not folder:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        p = ProgressPopup(self, t("processing"))
        self.progress_popup = p
        def logcb(m): p.log(m)
        # 直接主線程執行，避免COM跨線程錯誤
        OutlookDownloader.download_weekly_orders(folder, config_file=config, callback=logcb)

    def _run_enhanced_automation(self):
        source = self.folder_vars.get("source_folder", ctk.StringVar()).get()
        supplier = self.folder_vars.get("supplier_folder", ctk.StringVar()).get()
        master_config = self.master_config_var.get()
        if not source or not supplier or not master_config:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        p = ProgressPopup(self, t("processing"))
        self.progress_popup = p
        def logcb(m): p.log(m)
        config_mgr = UnifiedConfigManager(master_config)
        auto = EnhancedOrderAutomation(config_mgr)
        threading.Thread(target=lambda: p.log(auto.run_automation(source, supplier, log_callback=logcb)[1])).start()

    def _select_config_file(self, var=None, filetypes=None):
        f = filedialog.askopenfilename(title=t("select_folder"), filetypes=filetypes or [("Excel files", "*.xlsx;*.xls")])
        if f:
            if var:
                var.set(f)
            else:
                self.config_file_var.set(f)

    def _select_folder(self, key):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f:
            self.folder_vars[key].set(f)

    def _select_folder_var(self, var):
        f = filedialog.askdirectory(title=t("select_folder"))
        if f:
            var.set(f)

    def _show_required_outlets_window(self):
        import pandas as pd
        from tkinter import Toplevel, Scrollbar, VERTICAL, HORIZONTAL, RIGHT, BOTTOM, Y, X, BOTH
        import tkinter.ttk as ttk
        from tkinter import messagebox
        master_config = self.master_config_var.get()
        if not master_config:
            messagebox.showwarning("警告/Warning", "請先選擇統一配置文件/Master Config (Excel)！")
            return
        try:
            df_req = pd.read_excel(master_config, sheet_name="Supplier Requirements")
        except Exception as e:
            messagebox.showerror("錯誤/Error", f"讀取 Supplier Requirements 失敗: {e}")
            return
        ordered = {}
        for row in getattr(self, '_checklist_table_data', []):
            supplier = str(row["supplier"]).strip()
            outlet = str(row["outlet"]).strip()
            if supplier not in ordered:
                ordered[supplier] = set()
            ordered[supplier].add(outlet)
        def normalize(text):
            import re
            return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()
        win = ctk.CTkToplevel(self)
        win.title("必要門市清單/Required Outlets List")
        win.geometry("800x500")
        ctk.CTkLabel(win, text="📋 供應商/Supplier        ❗必要分店/Missing Outlets", font=FONT_BIGBTN).pack(pady=10)
        frame = ctk.CTkFrame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        style = ttk.Style()
        style.configure("Red.TLabel", foreground="#ef4444")
        tree = ttk.Treeview(frame, columns=("supplier", "missing"), show="headings", height=16)
        tree.heading("supplier", text="📋 供應商/Supplier")
        tree.heading("missing", text="❗必要分店/Missing Outlets")
        tree.column("supplier", width=180, anchor="center")
        tree.column("missing", width=600, anchor="w")
        # 先建立 config excel 的 outlet 別名 mapping
        outlet_name_to_short = {}
        if os.path.exists(master_config):
            try:
                df = pd.read_excel(master_config, sheet_name=None)
                outlet_df = None
                for key in df.keys():
                    if key.strip().lower() == "outlet":
                        outlet_df = df[key]
                        break
                if outlet_df is not None:
                    for _, row in outlet_df.iterrows():
                        short = str(row.get("Short Name", "")).strip()
                        full = str(row.get("Outlet Full Name", "")).strip()
                        email_name = str(row.get("Name in Email", "")).strip()
                        for n in [short, full, email_name]:
                            n_norm = normalize_name(n)
                            if n and n_norm:
                                outlet_name_to_short[n_norm] = short
            except Exception as e:
                print("Outlet mapping read error:", e)
        # checklist table 的每一筆 outlet 轉成 normalize 縮寫
        checklist_outlet_shorts = set()
        for s, outs in ordered.items():
            for o in outs:
                n = normalize_name(o)
                short = outlet_name_to_short.get(n, n)
                checklist_outlet_shorts.add(short)
        for _, row in df_req.iterrows():
            supplier = str(row.iloc[0]).strip()
            outlets = row.iloc[1]
            if pd.isna(supplier):
                continue
            if pd.isna(outlets) or not str(outlets).strip():
                continue
            required = set(o.strip() for o in str(outlets).split(",") if o.strip())
            # 找出缺少的分店，供應商名稱用 normalize 完全相等比對
            found = set()
            norm_supplier = normalize_name(supplier)
            for s, outs in ordered.items():
                norm_s = normalize_name(s)
                if norm_supplier == norm_s:
                    found = outs
                    break
            # 將 found 轉成 normalize 縮寫集合
            found_shorts = set()
            for f in found:
                n = normalize_name(f)
                short = outlet_name_to_short.get(n, n)
                found_shorts.add(short)
            missing = []
            for req in required:
                req_short = outlet_name_to_short.get(normalize_name(req), normalize_name(req))
                if req_short not in found_shorts:
                    missing.append(req)
            if missing:
                missing_list = [f"❗{o}" for o in missing]
                tree.insert("", "end", values=(supplier, ", ".join(missing_list)))
        tree.pack(fill="both", expand=True)
        vsb = Scrollbar(frame, orient=VERTICAL, command=tree.yview)
        vsb.pack(side=RIGHT, fill=Y)
        tree.configure(yscrollcommand=vsb.set)
        hsb = Scrollbar(frame, orient=HORIZONTAL, command=tree.xview)
        hsb.pack(side=BOTTOM, fill=X)
        tree.configure(xscrollcommand=hsb.set)
        def copy_selected():
            items = tree.selection()
            if not items:
                return
            import pyperclip
            rows = []
            for item in items:
                vals = tree.item(item, "values")
                rows.append("\t".join(vals))
            pyperclip.copy("\n".join(rows))
            messagebox.showinfo("複製成功/Copy Success", "已複製到剪貼簿！/Copied to clipboard!")
        ctk.CTkButton(win, text="複製所選/Copy Selected", command=copy_selected).pack(pady=5)

    def show_user_guide(self):
        """显示用户指南"""
        try:
            guide_path = resource_path("UserGuide.txt")
            if os.path.exists(guide_path):
                with open(guide_path, "r", encoding="utf-8") as f:
                    content = f.read()
                ScrollableMessageBox(self, "用户指南", content)
            else:
                messagebox.showwarning("指南缺失", "用户指南文件未找到")
        except Exception as e:
            messagebox.showerror("错误", f"无法加载用户指南: {str(e)}")

    def _run_cross_check_email_log(self):
        import os
        import re
        import pandas as pd
        import customtkinter as ctk
        folder = self.checklist_folder_var.get()
        table = getattr(self, '_checklist_table_data', [])
        config_path = self.master_config_var.get()
        email_log_path = os.path.join(folder, "email_bodies_log.txt")
        if not os.path.exists(email_log_path):
            from tkinter import messagebox
            messagebox.showinfo("交叉檢查結果\nCross Check Result", "找不到 email_bodies_log.txt！\nEmail log not found!")
            return
        # 1. 讀取 OUTLET mapping
        outlet_map = {}
        if os.path.exists(config_path):
            try:
                df = pd.read_excel(config_path, sheet_name=None)
                outlet_df = None
                for key in df.keys():
                    if key.strip().lower() == "outlet":
                        outlet_df = df[key]
                        break
                if outlet_df is not None:
                    for _, row in outlet_df.iterrows():
                        short = str(row.get("Short Name", "")).strip()
                        full = str(row.get("Outlet Full Name", "")).strip()
                        email = str(row.get("Email", "")).strip()
                        def norm(x):
                            return re.sub(r'[^a-zA-Z0-9]', '', x).lower()
                        if short:
                            outlet_map[norm(short)] = short
                        if full:
                            outlet_map[norm(full)] = short
                        if email:
                            outlet_map[norm(email)] = short
                # 額外：常見 outlet 別名自動加入
                for k in list(outlet_map.keys()):
                    if 'whitesand' in k:
                        outlet_map['ws'] = outlet_map[k]
                    if 'tampines' in k and 'smrt' in k:
                        outlet_map['tsmrt'] = outlet_map[k]
                    if 'junction8' in k or 'j8' in k:
                        outlet_map['j8'] = outlet_map[k]
            except Exception as e:
                print("Outlet mapping read error:", e)
        def normalize_name(name):
            if not name:
                return ""
            import re
            name = re.sub(r'\(.*?\)', '', name)
            name = re.sub(r'update', '', name, flags=re.IGNORECASE)
            name = re.sub(r'[^a-zA-Z0-9]', '', name)
            return name.strip().upper()
        def is_supplier_match(s1, s2):
            n1, n2 = normalize_name(s1), normalize_name(s2)
            return n1 in n2 or n2 in n1 or (len(n1) > 3 and n1[:4] in n2) or (len(n2) > 3 and n2[:4] in n1)
        def is_cover_status_ok(status):
            return any(x in status for x in ['✔', '✓'])
        # ===== Outlet mapping 與新版比對 function =====
        short_to_full = {}
        full_to_short = {}
        email_name_to_short = {}
        outlet_all_names = set()
        if os.path.exists(master_config):
            try:
                df = pd.read_excel(master_config, sheet_name=None)
                outlet_df = None
                for key in df.keys():
                    if key.strip().lower() == "outlet":
                        outlet_df = df[key]
                        break
                if outlet_df is not None:
                    for _, row in outlet_df.iterrows():
                        short = str(row.get("Short Name", "")).strip()
                        full = str(row.get("Outlet Full Name", "")).strip()
                        email_name = str(row.get("Name in Email", "")).strip()
                        n_short = normalize_name(short)
                        n_full = normalize_name(full)
                        n_email = normalize_name(email_name)
                        if short:
                            short_to_full[n_short] = full
                            outlet_all_names.add(n_short)
                        if full:
                            full_to_short[n_full] = short
                            outlet_all_names.add(n_full)
                        if email_name:
                            email_name_to_short[n_email] = short
                            outlet_all_names.add(n_email)
            except Exception as e:
                print("Outlet mapping read error:", e)
        def is_outlet_match(req_outlet, ordered_outlet):
            n_req = normalize_name(req_outlet)
            n_ordered = normalize_name(ordered_outlet)
            # 三者任兩個 normalize 後有 match 就 True
            if n_req and n_ordered and n_req == n_ordered:
                return True
            if n_req in outlet_all_names or n_ordered in outlet_all_names:
                return True
            for n in outlet_all_names:
                if n == n_req or n == n_ordered:
                    return True
            return False
        # 2. 解析 email log
        with open(email_log_path, "r", encoding="utf-8") as f:
            content = f.read()
        outlet_results = {}
        import re
        blocks = re.split(r"——— ?邮件 ?\d+ ?———", content)
        for block in blocks:
            lines = [l.strip() for l in block.strip().splitlines() if l.strip()]
            if not lines or len(lines) < 2:
                continue
            outlet_raw = ""
            for l in lines:
                m = re.match(r"\[发件人\](.+)", l)
                if m:
                    outlet_raw = m.group(1).strip()
                    break
            if not outlet_raw:
                continue
            outlet_norm = normalize_name(outlet_raw)
            outlet_short = outlet_map.get(outlet_norm.lower(), outlet_raw)
            # 先 map 成 short，再找 config excel 的全名
            outlet_full = outlet_raw
            if os.path.exists(config_path):
                try:
                    df = pd.read_excel(config_path, sheet_name=None)
                    outlet_df = None
                    for key in df.keys():
                        if key.strip().lower() == "outlet":
                            outlet_df = df[key]
                            break
                    if outlet_df is not None:
                        for _, row in outlet_df.iterrows():
                            short = str(row.get("Short Name", "")).strip()
                            full = str(row.get("Outlet Full Name", "")).strip()
                            if short and short == outlet_short:
                                outlet_full = full
                                break
                except Exception as e:
                    pass
            outlet_norm_full = normalize_name(outlet_full)
            claimed = []
            for line in lines:
                m = re.match(r"\d+\.\s*(.+)", line)
                if m:
                    claimed.append(m.group(1).strip())
            found = set()
            for row in table:
                row_outlet_norm = normalize_name(row["outlet"])
                # outlet 比對加強：允許 substring/前4碼
                outlet_match = is_outlet_match(row["outlet"], outlet_full)
                for s in claimed:
                    supplier_match = is_supplier_match(row["supplier"], s)
                    cover_ok = is_cover_status_ok(row["cover_status"])
                    print(f"[DEBUG] 比對: supplier: {normalize_name(row['supplier'])} <-> {normalize_name(s)} | outlet: {row['outlet']} <-> {outlet_full} | outlet_match: {outlet_match} | cover: {row['cover_status']} => {cover_ok}")
                    if outlet_full and outlet_match and supplier_match and cover_ok:
                        found.add(s)
            outlet_results[outlet_full] = []
            for supplier in claimed:
                if supplier in found:
                    outlet_results[outlet_full].append((supplier, True))
                else:
                    outlet_results[outlet_full].append((supplier, False))
        # 將 outlet_results 轉成表格資料
        table_data = []
        for outlet, suppliers in outlet_results.items():
            for supplier, ok in suppliers:
                status = "✔ 有整合！\nIntegrated!" if ok else "✗ 沒有整合！\nNot Integrated!"
                table_data.append((supplier, outlet, status, ok))
        class CrossCheckResultTable(ctk.CTkToplevel):
            def __init__(self, master, table_data):
                super().__init__(master)
                self.title("交叉檢查結果\nCross Check Result")
                self.geometry("700x700")
                self.search_var = ctk.StringVar()
                ctk.CTkLabel(self, text="搜尋/分店/供應商\nSearch Outlet/Supplier").pack(pady=(10, 0))
                search_entry = ctk.CTkEntry(self, textvariable=self.search_var)
                search_entry.pack(fill="x", padx=10)
                search_entry.bind("<KeyRelease>", self.update_filter)
                import tkinter as tk
                from tkinter import ttk
                frame = ctk.CTkFrame(self)
                frame.pack(fill="both", expand=True, padx=10, pady=10)
                columns = ("supplier", "outlet", "status")
                self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=25)
                self.tree.heading("supplier", text="供應商\nSupplier")
                self.tree.heading("outlet", text="分店\nOutlet")
                self.tree.heading("status", text="狀態\nStatus")
                self.tree.column("supplier", width=180)
                self.tree.column("outlet", width=120)
                self.tree.column("status", width=180)
                self.tree.pack(fill="both", expand=True)
                vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
                self.tree.configure(yscroll=vsb.set)
                vsb.pack(side="right", fill="y")
                self.full_data = [(s, o, st) for s, o, st, _ in table_data]
                self.update_table(self.full_data)
                ctk.CTkButton(self, text="關閉\nClose", command=self.destroy).pack(pady=10)
            def update_table(self, data):
                for row in self.tree.get_children():
                    self.tree.delete(row)
                for supplier, outlet, status in data:
                    self.tree.insert("", "end", values=(supplier, outlet, status))
            def update_filter(self, event=None):
                keyword = self.search_var.get().lower()
                filtered = [row for row in self.full_data if keyword in row[0].lower() or keyword in row[1].lower()]
                self.update_table(filtered)
        CrossCheckResultTable(self, table_data)

# ========== Outlook下载器 ==========
class OutlookDownloader:
    """Outlook订单下载器"""
    
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
    
    @staticmethod
    def get_smtp_address(msg):
        """取得郵件的 SMTP 格式寄件人"""
        try:
            if hasattr(msg, "SenderEmailType") and msg.SenderEmailType == "SMTP":
                return str(msg.SenderEmailAddress).lower().strip()
            if hasattr(msg, "Sender") and hasattr(msg.Sender, "GetExchangeUser"):
                ex_user = msg.Sender.GetExchangeUser()
                if ex_user and hasattr(ex_user, "PrimarySmtpAddress"):
                    return str(ex_user.PrimarySmtpAddress).lower().strip()
        except Exception:
            pass
        return ""

    @classmethod
    def read_outlet_config(cls, config_file):
        import pandas as pd
        df = pd.read_excel(config_file)
        # 自動偵測 short name/email 欄位名
        short_col = None
        email_col = None
        for col in df.columns:
            if 'short' in col.lower():
                short_col = col
            if 'email' in col.lower():
                email_col = col
        if not email_col:
            raise Exception("找不到 email 欄位")
        if not short_col:
            raise Exception("找不到 short name 欄位")
        outlets = []
        for _, row in df.iterrows():
            if pd.isna(row[email_col]) or pd.isna(row[short_col]) or not str(row[email_col]).strip() or not str(row[short_col]).strip():
                continue
            outlets.append({
                'email': clean_email(str(row[email_col])),
                'short_name': str(row[short_col]).strip(),
                # 其他欄位可依需求加
            })
        return outlets

    @classmethod
    def download_weekly_orders(cls, destination_folder, config_file=None, account_idx=None, callback=None):
        """下载周订单"""
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

        email_to_outlet = {}
        if config_file:
            try:
                outlets = cls.read_outlet_config(config_file)
                allowed_senders = set(clean_email(o['email']) for o in outlets if clean_email(o['email']))
                email_to_outlet = {clean_email(o['email']): o for o in outlets}
                if callback:
                    callback(f"✅ 已加载分店配置: {len(outlets)} 个分店")
            except Exception as e:
                if callback:
                    callback(f"❌ 分店配置读取失败: {str(e)}")
                return

        week_no = datetime.now().isocalendar()[1]
        save_path = os.path.join(destination_folder, f"Week_{week_no}")
        os.makedirs(save_path, exist_ok=True)
        cls.extracted_bodies = []

        today = datetime.now()
        monday_this_week = today - timedelta(days=today.weekday())
        start_of_range = monday_this_week - timedelta(days=2)
        end_of_range = monday_this_week + timedelta(days=6, hours=23, minutes=59)

        try:
            outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        except Exception as e:
            error_msg = f"無法啟動 Outlook: {str(e)}"
            if callback:
                callback(error_msg)
            return

        if account_idx is None:
            accounts = [outlook.Folders.Item(i + 1) for i in range(outlook.Folders.Count)]
            account_names = [acct.Name for acct in accounts]
            account_idx = simpledialog.askinteger(
                t("select_account"),
                "\n".join([f"[{i}] {name}" for i, name in enumerate(account_names)]) + "\n" + t("enter_index"),
                minvalue=0, maxvalue=len(account_names)-1,
                parent=None
            )
            if account_idx is None:
                return
    
        try:
            account_folder = outlook.Folders.Item(account_idx + 1)
            messages = cls._collect_messages(account_folder, start_of_range, end_of_range)
            debug_logs = []
            not_weekly_msgs = []
            not_downloaded_msgs = []
            latest_messages = {}
            allowed_senders = set(email_to_outlet.keys()) if email_to_outlet else set()
            for msg in messages:
                try:
                    sender_email = clean_email(cls.get_smtp_address(msg))
                    subject = (msg.Subject or "").lower()
                    # Debug log
                    print(f"allowed_senders: {allowed_senders}")
                    print(f"sender_email: {sender_email}")
                    # 只包含 weekly/order，排除 amendment/amend
                    is_weekly = ("weekly" in subject or "order" in subject) and ("amendment" not in subject and "amend" not in subject)
                    has_attachment = getattr(msg, "Attachments", None) and msg.Attachments.Count > 0
                    if allowed_senders and sender_email not in allowed_senders:
                        continue
                    if not is_weekly:
                        not_weekly_msgs.append(f"[非週訂單] {sender_email} | {msg.Subject}")
                        continue
                    if sender_email not in latest_messages or msg.ReceivedTime > latest_messages[sender_email].ReceivedTime:
                        latest_messages[sender_email] = msg
                    if not has_attachment:
                        not_downloaded_msgs.append(f"[無附件] {sender_email} | {msg.Subject}")
                except Exception as e:
                    debug_logs.append(f"[錯誤] {str(e)}")
            # 下載附件
            result = cls._download_attachments(list(latest_messages.values()), save_path, email_to_outlet, week_no)
            all_outlets = set(o['short_name'] for o in outlets) if config_file else set()
            downloaded_outlets = set(result['matched_outlets'])
            missing_outlets = all_outlets - downloaded_outlets
            summary = (
                f"=== {t('download_summary')} ===\n"
                f"📅 日期范围: {start_of_range.strftime('%Y-%m-%d')} 至 {end_of_range.strftime('%Y-%m-%d')}\n"
                f"✅ {t('auto_download')}: {result['downloaded']}\n"
                f"⏩ {t('skipped')}: {result['skipped']}\n"
                f"📁 {t('saved_to')}: {save_path}\n\n"
                f"已下載分店: {', '.join(sorted(downloaded_outlets)) if downloaded_outlets else '無'}\n"
                f"未下載分店: {', '.join(sorted(missing_outlets)) if missing_outlets else '無'}\n"
                f"匹配的分店: {len(result['matched_outlets'])}\n"
                f"未匹配的邮件: {len(result['unmatched_emails'])}"
            )
            if result['unmatched_emails']:
                summary += f"\n\n未匹配的邮箱:\n" + "\n".join(result['unmatched_emails'])
            if not_weekly_msgs:
                summary += f"\n\n[非週訂單郵件, 有附件但未下載]:\n" + "\n".join(not_weekly_msgs)
            if not_downloaded_msgs:
                summary += f"\n\n[有附件但未下載]:\n" + "\n".join(not_downloaded_msgs)
            if debug_logs:
                summary += f"\n\n[Debug]:\n" + "\n".join(debug_logs)
            if callback:
                callback(summary)
            # 下载后自动保存美化log
            def format_email_bodies(bodies):
                pretty = []
                for idx, b in enumerate(bodies, 1):
                    lines = b.split("\n")
                    sender = lines[0] if lines else ""
                    subject = lines[1] if len(lines) > 1 else ""
                    items = lines[2:] if len(lines) > 2 else []
                    if not items or all(not i.strip() for i in items):
                        items = ["无数字开头内容"]
                    pretty.append(f"——— 邮件 {idx} ———\n[发件人] {sender}\n[主题] {subject}\n[内容]")
                    for i, item in enumerate(items, 1):
                        pretty.append(f"  {i}. {item.strip()}")
                return "\n\n".join(pretty)
            if cls.extracted_bodies:
                log_file = os.path.join(save_path, "email_bodies_log.txt")
                pretty_text = format_email_bodies(cls.extracted_bodies)
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write(pretty_text)
        except Exception as e:
            error_msg = f"下載過程中出錯: {str(e)}"
            if callback:
                callback(error_msg)

    @classmethod
    def _collect_messages(cls, folder, start_date, end_date):
        """收集邮件"""
        messages = []
        try:
            filter_str = (
                f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %I:%M %p')}' AND "
                f"[ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %I:%M %p')}'"
            )
            
            items = folder.Items.Restrict(filter_str)
            items.Sort("[ReceivedTime]", True)
            messages.extend([msg for msg in items if msg.Class == 43])
        except Exception as e:
            print(f"Error accessing folder {folder.Name}: {e}")
        
        for sub in folder.Folders:
            messages.extend(cls._collect_messages(sub, start_date, end_date))
        
        return messages

    @classmethod
    def _filter_latest_messages(cls, messages, allowed_senders):
        latest_messages = {}
        print("allowed_senders:", repr(allowed_senders))  # debug
        for msg in messages:
            try:
                sender_email = clean_email(cls.get_smtp_address(msg))
                if allowed_senders and sender_email not in allowed_senders:
                    print(f"跳過寄件人: {repr(sender_email)}")
                    continue
                subject = (msg.Subject or "").lower()
                is_weekly = "weekly" in subject or "order" in subject
                print(f"Subject: {msg.Subject} | Sender: {repr(sender_email)} | is_weekly: {is_weekly} | matched: {sender_email in allowed_senders}")
                if not is_weekly:
                    continue
                if sender_email not in latest_messages or msg.ReceivedTime > latest_messages[sender_email].ReceivedTime:
                    latest_messages[sender_email] = msg
            except Exception as e:
                print(f"Error processing message: {e}")
        return list(latest_messages.values())

    @classmethod
    def _download_attachments(cls, messages, save_path, email_to_outlet=None, week_no=None):
        result = {
            "downloaded": 0,
            "skipped": 0,
            "matched_outlets": [],
            "unmatched_emails": set()
        }
        allowed_senders = set(email_to_outlet.keys()) if email_to_outlet else set()
        for msg in messages:
            try:
                sender_email = clean_email(cls.get_smtp_address(msg))
                if allowed_senders and sender_email not in allowed_senders:
                    print(f"跳過寄件人: {sender_email}")
                    result['skipped'] += 1
                    continue
                attachments = msg.Attachments
                print(f"處理寄件人: {sender_email}，主旨: {msg.Subject}，附件數: {attachments.Count}")
                if attachments.Count == 0:
                    continue
                outlet_info = email_to_outlet.get(sender_email)
                prefix = outlet_info['short_name'].strip().upper() if outlet_info and outlet_info['short_name'] else "UNKNOWN"
                result['matched_outlets'].append(prefix)
                for idx, att in enumerate(attachments, 1):
                    filename = att.FileName
                    ext = os.path.splitext(filename)[1]
                    new_filename = f"{prefix}_WeeklyOrder-Week{week_no}"
                    if attachments.Count > 1:
                        new_filename += f"_{idx}"
                    new_filename += ext
                    save_file = os.path.join(save_path, new_filename)
                    file_counter = 1
                    while os.path.exists(save_file):
                        save_file = os.path.join(save_path, f"{prefix}_WeeklyOrder-Week{week_no}_{file_counter}{ext}")
                        file_counter += 1
                    att.SaveAsFile(save_file)
                    print(f"Downloaded from: {sender_email} to {save_file}")
                    result['downloaded'] += 1
                # ========== 正文提取：只保留第一组连续数字开头内容 ==========
                import re
                body = getattr(msg, "Body", "") or ""
                subject = getattr(msg, "Subject", "") or ""
                sender = getattr(msg, "SenderName", "") or sender_email
                lines = body.splitlines()
                items = []
                started = False
                for line in lines:
                    line = line.strip()
                    if re.match(r"^\d+[\.、\)]?\s*[^\s].*", line):
                        started = True
                        m = re.match(r"^\d+[\.、\)]?\s*(.*)", line)
                        if m:
                            items.append(m.group(1).strip())
                    elif started:
                        break
                if not items:
                    items = ["无数字开头内容"]
                entry = f"{sender}\n{subject}\n" + "\n".join(items)
                cls.extracted_bodies.append(entry)
            except Exception as e:
                print(f"Error downloading attachment: {e}")
                result['unmatched_emails'].add(sender_email)
                result['skipped'] += 1
        return result

    extracted_bodies = []

# ========== 运营用品订单 ==========
class OperationSuppliesOrder:
    """运营用品订单处理"""
    
    @staticmethod
    def get_monthly_order_data(master_file):
        """获取月度订单数据"""
        try:
            wb = load_workbook(master_file, data_only=True)
            data_sheet = wb["Data"]
            
            outlets = []
            for row in data_sheet.iter_rows(min_row=2, max_col=7, values_only=True):
                if row[1] and row[2]:
                    outlets.append({
                        "brand": row[1],
                        "outlet": row[2],
                        "short_name": row[3],
                        "full_name": row[4],
                        "address": row[5],
                        "delivery_day": row[6]
                    })
            
            orders = defaultdict(dict)
            unit_prices = defaultdict(dict)
            
            for supplier in ["Freshening", "Legacy", "Unikleen"]:
                if supplier in wb.sheetnames:
                    ws = wb[supplier]
                    
                    if supplier == "Freshening":
                        unit_prices[supplier] = [ws[f'D{i}'].value for i in range(12, 46)]
                    elif supplier == "Legacy":
                        unit_prices[supplier] = [ws[f'D{i}'].value for i in range(12, 15)]
                    else:
                        unit_prices[supplier] = [ws[f'D{i}'].value for i in range(12, 30)]
            
            for outlet in outlets:
                sheet_name = outlet["short_name"]
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    
                    freshening = []
                    for row in range(4, 38):
                        cell_value = ws[f'L{row}'].value
                        freshening.append(cell_value if cell_value is not None else 0)
                    
                    legacy = []
                    for row in range(41, 44):
                        cell_value = ws[f'L{row}'].value
                        legacy.append(cell_value if cell_value is not None else 0)
                    
                    unikleen = []
                    for row in range(47, 65):
                        cell_value = ws[f'L{row}'].value
                        unikleen.append(cell_value if cell_value is not None else 0)
                    
                    orders[sheet_name] = {
                        "freshening": freshening,
                        "legacy": legacy,
                        "unikleen": unikleen
                    }
            
            templates = {}
            for supplier in ["Freshening", "Legacy", "Unikleen"]:
                if supplier in wb.sheetnames:
                    templates[supplier] = wb[supplier]
            
            return outlets, orders, templates, unit_prices
        except Exception as e:
            return None, f"Error reading master file: {str(e)}"

    @staticmethod
    def _is_number(val):
        try:
            float(val)
            return True
        except (TypeError, ValueError):
            return False

    @classmethod
    def calculate_order_amounts(cls, orders, unit_prices):
        """计算订单金额"""
        amounts = defaultdict(dict)
        for outlet, order_data in orders.items():
            for supplier in ["freshening", "legacy", "unikleen"]:
                items = order_data.get(supplier, [])
                prices = unit_prices.get(supplier.capitalize(), [])
                if len(prices) < len(items):
                    prices = prices + [0] * (len(items) - len(prices))
                total = sum(
                    float(qty) * float(price)
                    for qty, price in zip(items, prices)
                    if cls._is_number(qty) and cls._is_number(price)
                )
                amounts[outlet][supplier] = total
        return amounts

    @classmethod
    def check_moq(cls, outlets, orders, unit_prices, log_callback=None):
        """检查MOQ"""
        results = {
            "freshening": defaultdict(list),
            "legacy": defaultdict(list),
            "unikleen": defaultdict(list)
        }
        
        amounts = cls.calculate_order_amounts(orders, unit_prices)
        
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
        
        for outlet in outlets:
            short_name = outlet["short_name"]
            quantities = orders.get(short_name, {}).get("legacy", [])
            
            total = sum(q for q in quantities if isinstance(q, (int, float)))
            cartons = total
            
            if cartons < 2 and total > 0:
                results["legacy"]["below_moq"].append(f"{short_name} ({cartons} ctn < 2 ctn)")
            elif total > 0:
                results["legacy"]["above_moq"].append(f"{short_name} ({cartons} ctn)")
        
        for outlet in outlets:
            short_name = outlet["short_name"]
            amount = amounts.get(short_name, {}).get("unikleen", 0)
            
            if amount < 80 and amount > 0:
                results["unikleen"]["below_moq"].append(f"{short_name} (${amount:.2f} < $80)")
            elif amount > 0:
                results["unikleen"]["above_moq"].append(f"{short_name} (${amount:.2f})")
        
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
    def safe_set_cell_value(cls, ws, addr, value):
        from openpyxl.cell.cell import MergedCell
        cell = ws[addr]
        if not isinstance(cell, MergedCell):
            cell.value = value

    @classmethod
    def generate_supplier_files(
        cls, master_file, output_folder,
        outlets, orders, templates, amounts, log_callback=None
    ):
        """只訂一次貨：下月1號之後的第一個送貨日，且保證格式/公式/條件格式"""
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter
        import copy

        now = datetime.now()
        next_month = (now.month % 12) + 1
        year = now.year + (now.month // 12)
        saved = []

        for sup, tmpl_ws in templates.items():
            if log_callback: log_callback(f"\n--- Generating {sup} files ---")

            weekday_to_col = {
                'mon': 6, 'tue': 7, 'wed': 8, 'thu': 9, 'fri': 10, 'sat': 11, 'sun': 12
            }
            first_day = datetime(year, next_month, 1)
            first_wd_num = first_day.weekday()  # 0=Mon, 6=Sun
            first_wd = first_day.strftime("%a").lower()[:3]

            wb_out = Workbook()
            if wb_out.active is not None:
                wb_out.remove(wb_out.active)
            cnt = 0
            for o in outlets:
                sn = o["short_name"]
                data_list = orders.get(sn, {}).get(sup.lower(), [])
                nums = []
                for x in data_list:
                    try:
                        nums.append(float(x))
                    except:
                        nums.append(0.0)
                if not any(nums):
                    continue
                ws_new = wb_out.create_sheet(title=sn)
                # 複製模板內容
                for row in tmpl_ws.iter_rows():
                    for cell in row:
                        nc = ws_new.cell(row=cell.row, column=cell.column)
                        nc.value = cell.value
                        if cell.has_style:
                            nc.font = copy.copy(cell.font)
                            nc.border = copy.copy(cell.border)
                            nc.fill = copy.copy(cell.fill)
                            nc.number_format = cell.number_format
                            nc.protection = copy.copy(cell.protection)
                            nc.alignment = copy.copy(cell.alignment)
                if hasattr(tmpl_ws.conditional_formatting, '_cf_rules'):
                    for sqref, rules in tmpl_ws.conditional_formatting._cf_rules.items():
                        if isinstance(rules, list):
                            for rule in rules:
                                ws_new.conditional_formatting.add(sqref, rule)
                        else:
                            ws_new.conditional_formatting.add(sqref, rules)
                for ci in range(1, tmpl_ws.max_column+1):
                    cl = get_column_letter(ci)
                    ws_new.column_dimensions[cl].width = tmpl_ws.column_dimensions[cl].width
                for ri in range(1, tmpl_ws.max_row+1):
                    ws_new.row_dimensions[ri].height = tmpl_ws.row_dimensions[ri].height
                for mr in tmpl_ws.merged_cells.ranges:
                    ws_new.merge_cells(str(mr))
                # F5/F6
                cls.safe_set_cell_value(ws_new, "F5", o["full_name"])
                cls.safe_set_cell_value(ws_new, "F6", o["address"])
                # Freshening 寫入
                if sup == "Freshening":
                    cls.safe_set_cell_value(ws_new, "F8", o["delivery_day"])
                    # 解析所有送貨日
                    days = []
                    rawd = o["delivery_day"]
                    if isinstance(rawd, (int, float)):
                        d0 = datetime(1899, 12, 30) + timedelta(days=rawd)
                        days = [d0.strftime("%a").lower()[:3]]
                    else:
                        for p in str(rawd).split("/"):
                            if isinstance(p, str):
                                kd = p.strip().lower()[:3]
                            else:
                                kd = str(p).lower()[:3]
                            if kd in weekday_to_col:
                                days.append(kd)
                    days = list(dict.fromkeys(days))
                    # 找出下月1號之後的第一個送貨日
                    day_num_map = {'mon':0,'tue':1,'wed':2,'thu':3,'fri':4,'sat':5,'sun':6}
                    delivery_nums = [day_num_map[d] for d in days if d in day_num_map]
                    next_delivery = None
                    after = [d for d in delivery_nums if d >= first_wd_num]
                    if after:
                        next_delivery_num = min(after)
                    elif delivery_nums:
                        next_delivery_num = min(delivery_nums)
                    else:
                        next_delivery_num = None
                    # 反查 weekday_to_col
                    next_delivery_wd = None
                    for k, v in day_num_map.items():
                        if v == next_delivery_num:
                            next_delivery_wd = k
                            break
                    if next_delivery_wd and next_delivery_wd in weekday_to_col:
                        col = weekday_to_col[next_delivery_wd]
                        # 決定 row 起點
                        if first_wd in ['fri','sat','sun']:
                            base_row = 51
                        else:
                            base_row = 12
                        for row_idx, qty in enumerate(nums):
                            excel_row = base_row + row_idx
                            cls.safe_set_cell_value(ws_new, f"{get_column_letter(col)}{excel_row}", qty)
                elif sup == "Legacy":
                    # 只訂一次貨，找下月1號之後的第一個送貨日
                    days = ['mon','tue','wed','thu','fri']
                    day_num_map = {'mon':0,'tue':1,'wed':2,'thu':3,'fri':4}
                    delivery_nums = [day_num_map[d] for d in days if d in day_num_map]
                    after = [d for d in delivery_nums if d >= first_wd_num]
                    if after:
                        next_delivery_num = min(after)
                    elif delivery_nums:
                        next_delivery_num = min(delivery_nums)
                    else:
                        next_delivery_num = None
                    next_delivery_wd = None
                    for k, v in day_num_map.items():
                        if v == next_delivery_num:
                            next_delivery_wd = k
                            break
                    if next_delivery_wd and next_delivery_wd in weekday_to_col:
                        col = weekday_to_col[next_delivery_wd]
                        row = 12
                        for idx, qty in enumerate(nums):
                            cls.safe_set_cell_value(ws_new, f"{get_column_letter(col)}{row}", qty if idx == 0 else 0)
                elif sup == "Unikleen":
                    # 只訂一次貨，直接用下月1號是星期幾
                    col = weekday_to_col[first_wd]
                    row = 12
                    cls.safe_set_cell_value(ws_new, f"{get_column_letter(col)}{row}", nums[0] if len(nums) > 0 else 0)
                cnt += 1
            # 保存
            if cnt:
                fn = f"{sup}_Order_{year}_{next_month:02d}.xlsx"
                path = os.path.join(output_folder, fn)
                wb_out.save(path)
                saved.append(fn)
                if log_callback: log_callback(f"✅ Saved {fn} ({cnt} sheets)")
            else:
                if log_callback: log_callback(f"⚠️ {sup} 無訂單，未生成檔案")

        return True, saved

    @classmethod
    def process_order(cls, master_file, output_folder, log_callback=None, progress_callback=None):
        """处理订单"""
        try:
            if log_callback:
                log_callback(f"讀取主文件: {os.path.basename(master_file)}")
            outlets, orders, templates, unit_prices = cls.get_monthly_order_data(master_file)
            
            if not outlets:
                return False, "無法讀取分店數據，請檢查Data工作表"
            
            if log_callback:
                log_callback("\n計算訂購金額並檢查MOQ要求...")
            moq_results, moq_summary, amounts = cls.check_moq(outlets, orders, unit_prices, log_callback)
            
            if log_callback:
                log_callback("\n生成供應商訂單文件並顯示訂購金額...")
            success, supplier_files = cls.generate_supplier_files(
                master_file, output_folder, outlets, orders, templates, amounts, log_callback
            )
            
            if not success:
                return False, supplier_files
            
            result = (
                f"=== 營運用品月訂單處理完成 ===\n\n"
                f"📊 MOQ 檢查結果:\n{moq_summary}\n\n"
                f"📁 生成的供應商文件:\n" + 
                "\n".join([f"  - {file}" for file in supplier_files])
            )
            
            return True, result
        except Exception as e:
            return False, f"處理過程中出錯: {str(e)}\n{traceback.format_exc()}"

# ====== 在檔案開頭（import 之後）新增 ======
def clean_email(val):
    if not val:
        return ""
    import re
    # 去除所有空白字符（空格、Tab、换行、全角空格等），并转小写
    return re.sub(r"\s+", "", str(val)).lower().strip()

# ========== 支援公式自動計算的 cell 讀取（已棄用，改用 data_only=True） ==========
def get_cell_value_with_formula_support(ws, row, col, file_path=None, sheet_name=None):
    # 現在直接使用 cell.value，因為檔案已經用 data_only=True 讀取
    return ws.cell(row=row, column=col).value

# ========== 批量處理公式的優化版本 ==========
def process_formula_cells_batch(ws, file_path=None, sheet_name=None, max_rows=150):
    """批量處理所有公式 cell，一次啟動 Excel 處理整個檔案"""
    if not file_path or not sheet_name:
        return ws
    
    try:
        import xlwings as xw
        print(f"[DEBUG] 啟動 Excel 處理公式: {file_path}")
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        ws_xl = wb.sheets[sheet_name]
        
        # 掃描所有需要處理的 cell
        formula_cells = []
        for row in range(1, min(max_rows, ws.max_row + 1)):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                
                # 跳過合併儲存格
                if isinstance(cell, MergedCell):
                    continue
                
                val = cell.value
                
                # 檢查是否為公式（只處理真正的公式，不是 None）
                is_formula = (isinstance(val, str) and val.startswith("=")) or \
                            (hasattr(val, '__class__') and 'ArrayFormula' in str(val.__class__))
                
                if is_formula:
                    formula_cells.append((row, col))
        
        # 批量讀取所有公式 cell 的值
        if formula_cells:
            print(f"[DEBUG] 找到 {len(formula_cells)} 個公式 cell，開始批量處理")
            for row, col in formula_cells:
                try:
                    display_value = ws_xl.range((row, col)).value
                    # 更新 openpyxl 的 cell 值
                    ws.cell(row=row, column=col).value = display_value
                except Exception as e:
                    print(f"[DEBUG] 處理公式失敗 row={row} col={col}: {e}")
        
        wb.close()
        app.quit()
        print(f"[DEBUG] Excel 處理完成")
        
    except Exception as e:
        print(f"[DEBUG] 批量處理公式失敗: {e}")
    
    return ws

# ========== 入口点 ==========
if __name__ == '__main__':
    try:
        app = SushiExpressApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"Startup failed: {e}")
        sys.exit(1)

