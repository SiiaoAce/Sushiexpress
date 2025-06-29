
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
from PIL import Image, ImageTk, ImageDraw, ImageFont
from tkinter import filedialog, messagebox, simpledialog
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

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
VERSION = "2.7.0"
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
    """发光效果按钮"""
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
    """邮件发送确认对话框"""
    def __init__(self, parent, mail_item, supplier_name, outlet_name, attachment_path, on_confirm):
        super().__init__(parent)
        self.title(f"确认邮件 - {supplier_name}")
        self.geometry("800x600")
        self.transient(parent)
        self.grab_set()
        self.mail_item = mail_item
        self.on_confirm = on_confirm
        self.attachment_path = attachment_path
        
        self.configure(fg_color=DARK_BG)
        
        info_frame = ctk.CTkFrame(self, fg_color=DARK_PANEL, corner_radius=12)
        info_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkLabel(info_frame, text=f"收件人: {mail_item.To}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"主题: {mail_item.Subject}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"供应商: {supplier_name}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        ctk.CTkLabel(info_frame, text=f"分店: {outlet_name}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        
        if attachment_path:
            file_name = os.path.basename(attachment_path)
            ctk.CTkLabel(info_frame, text=f"附件: {file_name}", font=FONT_MID).pack(anchor="w", padx=10, pady=5)
        
        body_frame = ctk.CTkFrame(self, fg_color=DARK_PANEL, corner_radius=12)
        body_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        ctk.CTkLabel(body_frame, text="邮件正文:", font=FONT_BIGBTN).pack(anchor="w", padx=10, pady=5)
        self.body_text = ctk.CTkTextbox(body_frame, wrap="word", height=200, font=FONT_MID)
        self.body_text.pack(fill="both", expand=True, padx=10, pady=5)
        self.body_text.insert("1.0", mail_item.Body)
        
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=10)
        
        ctk.CTkButton(
            btn_frame, 
            text="发送邮件", 
            command=self._send_email,
            fg_color=ACCENT_GREEN,
            hover_color=BTN_HOVER
        ).pack(side="right", padx=10)
        
        ctk.CTkButton(
            btn_frame, 
            text="取消", 
            command=self.destroy
        ).pack(side="right", padx=10)
        
        ctk.CTkButton(
            btn_frame, 
            text="编辑正文", 
            command=self._edit_body,
            fg_color=ACCENT_BLUE,
            hover_color=BTN_HOVER
        ).pack(side="left", padx=10)
    
    def _edit_body(self):
        self.body_text.configure(state="normal")
    
    def _send_email(self):
        try:
            self.mail_item.Body = self.body_text.get("1.0", "end-1c")
            self.mail_item.Send()
            self.on_confirm(True, "邮件发送成功！")
            self.destroy()
        except Exception as e:
            self.on_confirm(False, f"邮件发送失败: {str(e)}")

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
                    supplier = row['supplier'].strip()
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
        """获取特定供应商-分店的送货日"""
        outlet_specific = self.schedule.get(supplier, {}).get(outlet_code)
        if outlet_specific is not None:
            return outlet_specific
        return self.schedule.get(supplier, {}).get('*', set())

    def validate_order(self, supplier, outlet_code, order_date, log_callback=None):
        """验证订单日期"""
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
                            "code": row[0].strip().upper(),
                            "name": row[1].strip() if len(row) > 1 else "",
                            "email": row[2].strip().lower() if len(row) > 2 else "",
                            "address": row[3].strip() if len(row) > 3 else "",
                            "delivery_day": row[4].strip() if len(row) > 4 else "",
                            "brand": row[5].strip() if len(row) > 5 else "Dine-In"
                        })
            
            # 读取供应商信息
            if "Suppliers" in wb.sheetnames:
                ws = wb["Suppliers"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        self.suppliers.append({
                            "name": row[0].strip(),
                            "email": row[1].strip().lower() if len(row) > 1 else "",
                            "contact": row[2].strip() if len(row) > 2 else "",
                            "phone": row[3].strip() if len(row) > 3 else ""
                        })
            
            # 读取配送日程
            if "Delivery Schedule" in wb.sheetnames:
                ws = wb["Delivery Schedule"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        self.delivery_schedule.append({
                            "supplier": row[0].strip(),
                            "outlet_code": row[1].strip().upper() if row[1] != "ALL" else "ALL",
                            "delivery_days": row[2].strip()
                        })
            
            # 读取邮件模板
            if "Email Templates" in wb.sheetnames:
                ws = wb["Email Templates"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        self.email_templates[row[0].strip()] = {
                            "subject": row[1].strip() if len(row) > 1 else "",
                            "body": row[2].strip() if len(row) > 2 else ""
                        }
            
            # 读取供应商要求
            if "Supplier Requirements" in wb.sheetnames:
                ws = wb["Supplier Requirements"]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and row[0]:
                        supplier_name = row[0].strip()
                        outlet_codes = [code.strip().upper() for code in str(row[1]).split(",") if code.strip()]
                        
                        if supplier_name not in self.supplier_requirements:
                            self.supplier_requirements[supplier_name] = []
                        
                        self.supplier_requirements[supplier_name] = outlet_codes
            
            return True, f"成功加载配置文件: {len(self.outlets)} 分店, {len(self.suppliers)} 供应商"
        except Exception as e:
            return False, f"加载配置文件失败: {str(e)}"
    
    def get_outlet(self, code):
        """根据分店代码获取分店信息"""
        for outlet in self.outlets:
            if outlet["code"] == code:
                return outlet
        return None
    
    def get_supplier(self, name):
        """根据供应商名称获取供应商信息"""
        for supplier in self.suppliers:
            if supplier["name"] == name:
                return supplier
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
    
    def send_email(self, recipient, supplier_name, body, attachment_path=None):
        """使用Outlook发送邮件"""
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            mail.To = recipient
            mail.Subject = self._get_standard_subject(supplier_name)
            mail.Body = body
            
            if attachment_path and os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
            
            return mail
        except Exception as e:
            return None, f"创建邮件失败: {str(e)}"

# ========== 订单自动化核心 ==========
class OrderAutomation:
    """订单自动化工具"""
    
    OUTLET_LIST = [
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
        ("Sushi TakeOut CityVibe (GTM)", "GTM"),
        ("Sushi TakeOut Tampines SMRT", "TSMRT"),
        ("Sushi TakeOut Woodlands", "WL"),
        ("Sushi TakeOut Toa Payoh", "TPY"),
        ("Sushi GOGO Junction 8", "J8"),
        ("Sushi GOGO Hougang MRT", "HGTO"),
        ("Sushi GOGO Oasis Terraces", "OASIS"),
        ("Sushi GOGO Sengkang MRT", "SKMRT"),
        ("Sushi GOGO Yew Tee Square", "YTS"),
        ("Sushi GOGO The Poiz Centre", "TPC"),
        ("Sushi GOGO Ang Mo Kio Hub", "AMK"),
        ("Sushi GOGO Canberra Plaza", "CBP"),
        ("Sushi GOGO Bukit Gombak", "BGM"),
        ("Sushi GOGO Pasir Ris Mall", "PRM"),
        ("Sushi GOGO North Point City", "NPG"),
        ("SushiPlus Bugis", "Bugis"),
        ("SushiPlus 313 Somerset", "313"),
        ("SushiPlus Tampines One", "T1"),
    ]

    @staticmethod
    def get_short_code(f5val):
        """获取分店简称"""
        val = (str(f5val) or "").strip().lower()
        
        for full, code in OrderAutomation.OUTLET_LIST:
            if full.lower() == val:
                return code
        
        for full, code in OrderAutomation.OUTLET_LIST:
            if full.lower() in val:
                return code
        
        for full, code in OrderAutomation.OUTLET_LIST:
            key_words = [x.lower() for x in full.split() if len(x) > 1]
            if key_words and all(kw in val for kw in key_words[-2:]):
                return code
        
        for full, code in OrderAutomation.OUTLET_LIST:
            if code.lower() in val:
                return code
        
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
        
        code = "".join([x for x in val if x.isalnum()])
        return code[:4].upper() if code else "UNKNOWN"

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
    def find_delivery_date_row(cls, ws, next_week_start, next_week_end, max_rows=150):
        """查找送货日期行"""
        valid_col_range = range(5, 12)
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
                label_cell = ws.cell(row=row_idx, column=5)
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
        """运行订单自动化"""
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
        
        supplier_to_outlets = defaultdict(list)
        outlet_to_suppliers = defaultdict(list)
        
        os.makedirs(supplier_folder, exist_ok=True)
        os.makedirs(outlet_folder, exist_ok=True)
        
        def log(message, include_timestamp=True):
            timestamp = datetime.now().strftime("%H:%M:%S") if include_timestamp else ""
            log_line = f"[{timestamp}] {message}" if include_timestamp else message
            log_lines.append(log_line + "\n")
            if log_callback:
                log_callback(log_line + "\n")
        
        def update_mapping():
            if mapping_callback:
                mapping_text = "分店-供应商对应关系:\n"
                for outlet, suppliers in outlet_to_suppliers.items():
                    mapping_text += f"{outlet}: {', '.join(suppliers)}\n"
                mapping_callback(mapping_text)
        
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
                    
                    f5_value = ws['F5'].value if 'F5' in ws else ws.cell(row=5, column=6).value
                    outlet_short = cls.get_short_code(f5_value)
                    log(f"  Sheet: {sheetname}, F5 Value: '{f5_value}', Short Code: {outlet_short}")
                    
                    if outlet_mapping:
                        outlet_info = outlet_mapping.get(outlet_short)
                        if outlet_info:
                            sheet_address = ws['F6'].value if 'F6' in ws else None
                            if sheet_address:
                                def normalize(addr):
                                    return re.sub(r'[^a-zA-Z0-9]', '', str(addr).lower())
                                
                                config_addr = normalize(outlet_info.get('address', ''))
                                sheet_addr = normalize(sheet_address)
                                
                                if config_addr != sheet_addr:
                                    log(f"  ❌ 地址不匹配: {outlet_short}")
                                    log(f"    配置地址: {outlet_info.get('address', '')}")
                                    log(f"    工作表地址: {sheet_address}")
                                    outlet_short = f"{outlet_short} (地址错误)"
                    
                    if date_validator:
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
                    
                    header_row, target_cols = cls.find_delivery_date_row(ws, next_week_start, next_week_end)
                    
                    has_order = False
                    if target_cols:
                        log(f"    Found delivery dates at row {header_row}, columns {[get_column_letter(c) for c in target_cols]}")
                        
                        for row_idx in range(header_row + 1, header_row + 100):
                            label_cell = ws.cell(row=row_idx, column=5)
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
        
        log("\nOutlet-Supplier Mapping:")
        mapping_text = ""
        for outlet, suppliers in outlet_to_suppliers.items():
            log(f"  {outlet}: {', '.join(suppliers)}")
            mapping_text += f"{outlet}: {', '.join(suppliers)}\n"
        
        if mapping_callback:
            mapping_callback(mapping_text)

        log("\nSupplier-Outlet Mapping:")
        for supplier, outlets in supplier_to_outlets.items():
            outlet_list = [o[0] for o in outlets]
            log(f"  {supplier}: {', '.join(outlet_list)}")

        log("\nCreating supplier files...")
        supplier_files = []
        for sheetname, outlet_file_pairs in supplier_to_outlets.items():
            supplier_path = os.path.join(supplier_folder, f"{sheetname}_Week_{next_week_start.strftime('%V')}.xlsx")
            new_wb = Workbook()
            new_wb.remove(new_wb.active)
            
            log(f"Creating supplier file for {sheetname} at {supplier_path}")
            
            for outlet, src_file, original_sheet in outlet_file_pairs:
                try:
                    log(f"  Adding outlet: {outlet} from {src_file}")
                    src_wb = load_workbook(src_file, data_only=False)
                    src_ws = src_wb[original_sheet]
                    
                    if src_ws.sheet_state != "visible":
                        log(f"    ⏩ Skipping hidden sheet: {original_sheet}")
                        continue
                    
                    target_ws = new_wb.create_sheet(title=outlet)
                    
                    for row in src_ws.iter_rows():
                        for cell in row:
                            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                            
                            if cell.has_style:
                                new_cell.font = copy.copy(cell.font)
                                new_cell.border = copy.copy(cell.border)
                                new_cell.fill = copy.copy(cell.fill)
                                new_cell.number_format = cell.number_format
                                new_cell.protection = copy.copy(cell.protection)
                                new_cell.alignment = copy.copy(cell.alignment)
                    
                    for col_idx in range(1, src_ws.max_column + 1):
                        col_letter = get_column_letter(col_idx)
                        target_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                    
                    for row_idx in range(1, src_ws.max_row + 1):
                        target_ws.row_dimensions[row_idx].height = src_ws.row_dimensions[row_idx].height
                    
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

        log("\nCreating outlet files...")
        outlet_files = []
        outlet_to_sheets = defaultdict(list)
        for sheetname, pairs in supplier_to_outlets.items():
            for outlet, path, original_sheet in pairs:
                outlet_to_sheets[outlet].append((sheetname, path, original_sheet))

        for outlet, supplier_sheets in outlet_to_sheets.items():
            out_path = os.path.join(outlet_folder, f"{outlet}_Week_{next_week_start.strftime('%V')}.xlsx")
            out_wb = Workbook()
            out_wb.remove(out_wb.active)
            
            log(f"Creating outlet file for {outlet} at {out_path}")
            
            for supplier, src_file, original_sheet in supplier_sheets:
                try:
                    log(f"  Adding supplier: {supplier} from {src_file}")
                    src_wb = load_workbook(src_file, data_only=False)
                    src_ws = src_wb[original_sheet]
                    
                    sheet_title = supplier[:31]
                    target_ws = out_wb.create_sheet(title=sheet_title)
                    
                    for row in src_ws.iter_rows():
                        for cell in row:
                            new_cell = target_ws.cell(
                                row=cell.row, 
                                column=cell.column, 
                                value=cell.value
                            )
                            
                            if cell.has_style:
                                new_cell.font = copy.copy(cell.font)
                                new_cell.border = copy.copy(cell.border)
                                new_cell.fill = copy.copy(cell.fill)
                                new_cell.number_format = cell.number_format
                                new_cell.protection = copy.copy(cell.protection)
                                new_cell.alignment = copy.copy(cell.alignment)
                    
                    for col_idx in range(1, src_ws.max_column + 1):
                        col_letter = get_column_letter(col_idx)
                        target_ws.column_dimensions[col_letter].width = src_ws.column_dimensions[col_letter].width
                    
                    for row_idx in range(1, src_ws.max_row + 1):
                        target_ws.row_dimensions[row_idx].height = src_ws.row_dimensions[row_idx].height
                    
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

        result_text = "\n订单整合结果:\n"
        result_text += f"已处理供应商文件: {len(supplier_files)}\n"
        result_text += f"已处理分店文件: {len(outlet_files)}\n\n"
        
        result_text += "供应商文件列表:\n" + "\n".join([f"- {file}" for file in supplier_files]) + "\n\n"
        result_text += "分店文件列表:\n" + "\n".join([f"- {file}" for file in outlet_files])
        
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

    def run_checklist(self, folder, delivery_config=None, log_callback=None):
        """运行检查表"""
        date_validator = DeliveryDateValidator(delivery_config) if delivery_config else None
        
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
            normalized_files = {self._normalize(f): f for f in files}
            output = []
            
            for supplier, keywords in supplier_keywords.items():
                match = self._find_supplier_file(normalized_files, keywords)
                if not match:
                    output.append(f"\n❌ {supplier} - Supplier file not found.")
                    continue
                
                result = self._process_supplier_file(
                    folder, match, must_have_outlets[supplier], 
                    date_validator, log_callback
                )
                output.extend(result)
            
            return "\n".join(output)
        except Exception as e:
            return f"❌ Error running checklist: {str(e)}\n{traceback.format_exc()}"

    @staticmethod
    def _normalize(text):
        """标准化文本"""
        return re.sub(r'[\s\(\)]', '', text.lower())

    @classmethod
    def _find_supplier_file(cls, normalized_files, keywords):
        """查找供应商文件"""
        for k in keywords:
            nk = cls._normalize(k)
            for nf, of in normalized_files.items():
                if nk in nf:
                    return of
        return None

    @classmethod
    def _process_supplier_file(cls, folder, filename, required_outlets, date_validator=None, log_callback=None):
        """处理供应商文件"""
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

# ========== 增强版订单自动化 ==========
class EnhancedOrderAutomation(OrderAutomation):
    """支持邮件发送的订单自动化"""
    
    def __init__(self, config_manager=None):
        super().__init__()
        self.config_manager = config_manager
        self.email_sender = EmailSender(config_manager)
    
    def run_automation(self, source_folder, supplier_folder, outlet_folder, 
                      log_callback=None, mapping_callback=None, 
                      email_callback=None):
        """运行自动化流程"""
        success, result = super().run_automation(
            source_folder, supplier_folder, outlet_folder,
            log_callback, mapping_callback
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
        self.geometry("1200x900")
        self.minsize(1000,800)
        self.configure(fg_color=DARK_BG)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # 初始化变量
        self.download_folder_var = None
        self.checklist_folder_var = None
        self.folder_vars = {}
        self.master_file_var = None
        self.output_folder_var = None
        self.progress_popup = None
        self.mapping_popup = None
        self.outlet_config_var = None
        self.delivery_config_var = None
        self.email_dialogs = []

        self._setup_ui()
        self.show_login()

    def _setup_ui(self):
        """设置UI布局"""
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # 头部
        self.header_frame = ctk.CTkFrame(self.main_container, fg_color=DARK_PANEL, corner_radius=24, height=150)
        self.header_frame.pack(fill="x", padx=10, pady=10)
        
        self.logo_img = load_image(LOGO_PATH)
        if self.logo_img:
            ctk.CTkLabel(self.header_frame, image=self.logo_img, text="").pack(expand=True)
        
        self.title_label = ctk.CTkLabel(self.header_frame, text="", font=FONT_TITLE, text_color=ACCENT_BLUE)
        self.subtitle_label = ctk.CTkLabel(self.header_frame, text="", font=FONT_SUB, text_color=TEXT_COLOR)
        self.title_label.pack()
        self.subtitle_label.pack()

        # 按钮区域
        self.buttons_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.buttons_frame.pack(fill="x", padx=10, pady=10)

        # 内容区域
        self.content_frame = ctk.CTkFrame(self.main_container, fg_color=DARK_PANEL, corner_radius=24)
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.content_frame.pack_propagate(False)

    def _on_close(self):
        """关闭应用程序确认"""
        if messagebox.askyesno(t("exit_system"), t("exit_confirm")):
            self.destroy()

    def show_login(self):
        """显示登录界面"""
        for w in self.content_frame.winfo_children(): w.destroy()
        self.login_frame = ctk.CTkFrame(self.content_frame, fg_color=DARK_PANEL, corner_radius=24, width=450, height=400)
        self.login_frame.place(relx=0.5, rely=0.5, anchor="center")
        self.pwd_entry = ctk.CTkEntry(self.login_frame, show="*", font=FONT_SUB, width=300, placeholder_text=t("password"), fg_color=ENTRY_BG, height=45)
        self.pwd_entry.pack(pady=(20,10))
        self.pwd_entry.bind("<Return>", lambda e: self._try_login())
        ctk.CTkButton(self.login_frame, text=t("login_btn"), command=self._try_login, width=200).pack(pady=(0,20))
        ctk.CTkLabel(self.login_frame, text=f"Version {VERSION} | {DEVELOPER}", font=("Consolas",12), text_color="#64748B").pack(side="bottom", pady=10)

    def _try_login(self):
        """尝试登录"""
        if self.pwd_entry.get() == PASSWORD:
            self.login_frame.destroy()
            self.show_main_menu()
        else:
            messagebox.showerror(t("error"), t("incorrect_pw"))

    def show_main_menu(self):
        """显示主菜单"""
        for w in self.content_frame.winfo_children(): w.destroy()
        self.title_label.configure(text=t("main_title"))
        self.subtitle_label.configure(text=t("select_function"))
        
        for w in self.buttons_frame.winfo_children(): w.destroy()
        
        funcs = [ (t("download_title"), self.show_download_ui),
                  (t("checklist_title"), self.show_checklist_ui),
                  (t("automation_title"), self.show_automation_ui),
                  ("營運用品月訂單", self.show_operation_supplies_ui),
                  (t("exit_system"), self._on_close) ]
        
        for text, cmd in funcs:
            btn = ctk.CTkButton(self.buttons_frame, text=text, command=cmd)
            btn.pack(side="left", padx=10, pady=10, expand=True)

    def show_function_ui(self, title, subtitle, content_callback):
        """显示功能界面"""
        for w in self.content_frame.winfo_children(): w.destroy()
        self.title_label.configure(text=title)
        self.subtitle_label.configure(text=subtitle)
        content = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=20, pady=20)
        content_callback(content)
        ctk.CTkButton(self.content_frame, text=t("back_to_menu"), command=self.show_main_menu).pack(pady=10)

    def show_download_ui(self):
        """显示下载界面"""
        def build(c):
            self.download_folder_var = ctk.StringVar()
            self.config_file_var = ctk.StringVar()
            
            row1 = ctk.CTkFrame(c)
            row1.pack(fill="x", pady=5)
            ctk.CTkLabel(row1, text="下载文件夹:").pack(side="left")
            ctk.CTkEntry(row1, textvariable=self.download_folder_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=t("browse"), command=self._select_download_folder).pack(side="left")
            
            row2 = ctk.CTkFrame(c)
            row2.pack(fill="x", pady=5)
            ctk.CTkLabel(row2, text="分店配置文件 (可选):").pack(side="left")
            ctk.CTkEntry(row2, textvariable=self.config_file_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=t("browse"), command=self._select_config_file).pack(side="left")
            
            ctk.CTkButton(c, text=t("start_download"), command=self._run_download).pack(pady=20)
        
        self.show_function_ui(t("download_title"), t("download_desc"), build)

    def _select_download_folder(self):
        """选择下载文件夹"""
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.download_folder_var.set(f)
    
    def _select_config_file(self):
        """选择配置文件"""
        f = filedialog.askopenfilename(
            title="选择分店配置文件",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if f: 
            self.config_file_var.set(f)

    def _run_download(self):
        """运行下载"""
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
        """显示检查表界面"""
        def build(c):
            self.checklist_folder_var = ctk.StringVar()
            self.delivery_config_var = ctk.StringVar()
            self.master_config_var = ctk.StringVar()
            
            row1 = ctk.CTkFrame(c)
            row1.pack(fill="x", pady=5)
            ctk.CTkLabel(row1, text="订单文件夹:").pack(side="left")
            ctk.CTkEntry(row1, textvariable=self.checklist_folder_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row1, text=t("browse"), command=self._select_checklist_folder).pack(side="left")
            
            row2 = ctk.CTkFrame(c)
            row2.pack(fill="x", pady=5)
            ctk.CTkLabel(row2, text="送货日期配置 (可选):").pack(side="left")
            ctk.CTkEntry(row2, textvariable=self.delivery_config_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row2, text=t("browse"), command=self._select_delivery_config).pack(side="left")
            
            row3 = ctk.CTkFrame(c)
            row3.pack(fill="x", pady=5)
            ctk.CTkLabel(row3, text="统一配置文件 (Excel):").pack(side="left")
            ctk.CTkEntry(row3, textvariable=self.master_config_var, state="readonly").pack(side="left", expand=True, fill="x", padx=5)
            ctk.CTkButton(row3, text=t("browse"), command=lambda: self._select_config_file(self.master_config_var, [("Excel files", "*.xlsx")])).pack(side="left")
            
            ctk.CTkButton(c, text=t("run_check"), command=self._run_enhanced_checklist).pack(pady=20)
        
        self.show_function_ui(t("checklist_title"), t("checklist_desc"), build)

    def _select_checklist_folder(self):
        """选择检查表文件夹"""
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.checklist_folder_var.set(f)
    
    def _select_delivery_config(self):
        """选择送货配置"""
        f = filedialog.askopenfilename(
            title="选择送货日期配置",
            filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if f: 
            self.delivery_config_var.set(f)

    def _run_enhanced_checklist(self):
        """运行增强版检查表"""
        folder = self.checklist_folder_var.get()
        if not folder:
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        
        config_manager = None
        if self.master_config_var.get():
            config_manager = UnifiedConfigManager(self.master_config_var.get())
            success, msg = config_manager.load_config(self.master_config_var.get())
            if not success:
                messagebox.showerror("配置错误", msg)
                return
        
        delivery_config = self.delivery_config_var.get() or None
        
        p = ProgressPopup(self, t("check_results"))
        self.progress_popup = p
        
        def logcb(m): p.log(m)
        
        threading.Thread(target=lambda: 
            self._thread_task(
                lambda: EnhancedOrderChecker(config_manager).run_checklist(folder, delivery_config, logcb),
                show_message=False
            )
        ).start()

    def show_automation_ui(self):
        """显示自动化界面"""
        def build(c):
            self.folder_vars = {}
            self.master_config_var = ctk.StringVar()
            
            for label, key in [(t("source_folder"),"source"),(t("supplier_folder"),"supplier"),(t("outlet_folder"),"outlet")]:
                row = ctk.CTkFrame(c); row.pack(pady=5)
                var = ctk.StringVar(); self.folder_vars[key] = var
                ctk.CTkLabel(row, text=label).pack(side="left", padx=5)
                ctk.CTkEntry(row, textvariable=var, state="readonly").pack(side="left", padx=5)
                ctk.CTkButton(row, text=t("browse"), command=lambda k=key: self._select_folder(k)).pack(side="left", padx=5)
            
            row = ctk.CTkFrame(c); row.pack(pady=5)
            ctk.CTkLabel(row, text="统一配置文件 (Excel):").pack(side="left", padx=5)
            ctk.CTkEntry(row, textvariable=self.master_config_var, state="readonly").pack(side="left", expand=True, padx=5)
            ctk.CTkButton(row, text=t("browse"), command=lambda: self._select_config_file(self.master_config_var, [("Excel files", "*.xlsx")])).pack(side="left", padx=5)
            
            ctk.CTkButton(c, text=t("start_automation"), command=self._run_enhanced_automation).pack(pady=20)
        
        self.show_function_ui(t("automation_title"), t("automation_desc"), build)

    def _select_folder(self, key):
        """选择文件夹"""
        f = filedialog.askdirectory(title=t("select_folder"))
        if f: self.folder_vars[key].set(f)
    
    def _select_config_file(self, var, filetypes=None):
        """选择配置文件"""
        f = filedialog.askopenfilename(
            title="选择配置文件",
            filetypes=filetypes or [("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if f: 
            var.set(f)

    def _run_enhanced_automation(self):
        """运行增强版自动化"""
        src, sup, out = self.folder_vars.get('source'), self.folder_vars.get('supplier'), self.folder_vars.get('outlet')
        if not all([src.get(), sup.get(), out.get()]): 
            messagebox.showwarning(t("warning"), t("folder_warning"))
            return
        
        config_manager = None
        if self.master_config_var.get():
            config_manager = UnifiedConfigManager(self.master_config_var.get())
            success, msg = config_manager.load_config(self.master_config_var.get())
            if not success:
                messagebox.showerror("配置错误", msg)
                return
        
        p = ProgressPopup(self, t("processing_orders"))
        self.progress_popup = p
        self.email_dialogs = []
        
        def logcb(m): p.log(m)
        def mapcb(m): 
            mp = MappingPopup(self, t("outlet_suppliers")); mp.update_mapping(m)
            self.mapping_popup = mp
        
        def emailcb(supplier_files):
            self.after(0, lambda: self._prepare_email_sending(supplier_files, config_manager))
        
        automation = EnhancedOrderAutomation(config_manager)
        
        threading.Thread(target=lambda: self._thread_task(
            lambda: automation.run_automation(
                src.get(), sup.get(), out.get(), 
                logcb, mapcb, emailcb)
        )).start()
    
    def _prepare_email_sending(self, supplier_files, config_manager):
        """准备发送邮件"""
        if not supplier_files:
            messagebox.showinfo("无文件可发送", "没有生成供应商文件")
            return
        
        for file_info in supplier_files:
            supplier_name = file_info["supplier"]
            supplier_info = config_manager.get_supplier(supplier_name) if config_manager else None
            
            if not supplier_info or not supplier_info.get("email"):
                self.progress_popup.log(f"⚠️ 跳过 {supplier_name} - 无邮箱配置")
                continue
            
            body = f"Dear {supplier_name},\n\nPlease find attached the weekly order for your reference.\n\nBest regards,\nSushi Express Purchasing Team"
            
            mail_item = self.email_sender.send_email(
                supplier_info["email"],
                supplier_name,
                body,
                file_info["path"]
            )
            
            if mail_item:
                dialog = EmailConfirmationDialog(
                    self,
                    mail_item,
                    supplier_name,
                    "",
                    file_info["path"],
                    self._email_confirmation_callback
                )
                self.email_dialogs.append(dialog)
            else:
                self.progress_popup.log(f"❌ 创建 {supplier_name} 邮件失败")
    
    def _email_confirmation_callback(self, success, message):
        """邮件发送确认回调"""
        if success:
            self.progress_popup.log(f"✅ {message}")
        else:
            self.progress_popup.log(f"❌ {message}")

    def show_operation_supplies_ui(self):
        """显示运营用品界面"""
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
    
    @classmethod
    def read_outlet_config(cls, config_file):
        """读取分店配置"""
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
                email_to_outlet = {o['email']: o for o in outlets}
                if callback:
                    callback(f"✅ 已加载分店配置: {len(outlets)} 个分店")
            except Exception as e:
                if callback:
                    callback(f"❌ 分店配置读取失败: {str(e)}")
                return

        week_no = datetime.now().isocalendar()[1]
        save_path = os.path.join(destination_folder, f"Week_{week_no}")
        os.makedirs(save_path, exist_ok=True)

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
    def _filter_latest_messages(cls, messages, email_to_outlet=None):
        """过滤最新邮件"""
        latest_messages = {}
        
        for msg in messages:
            try:
                sender_email = msg.SenderEmailAddress.lower()
                
                outlet_key = sender_email
                if email_to_outlet:
                    outlet_info = email_to_outlet.get(sender_email)
                    
                    if not outlet_info:
                        sender_name = msg.SenderName.lower()
                        for email, info in email_to_outlet.items():
                            if info['full_name'].lower() in sender_name:
                                outlet_info = info
                                outlet_key = info['short_name'].lower()
                                break
                
                subject = (msg.Subject or "").lower()
                is_weekly = "weekly" in subject or "order" in subject
                
                if not is_weekly:
                    continue
                
                if outlet_key not in latest_messages or msg.ReceivedTime > latest_messages[outlet_key].ReceivedTime:
                    latest_messages[outlet_key] = msg
                    
            except Exception as e:
                print(f"Error processing message: {e}")
        
        return list(latest_messages.values())

    @classmethod
    def _download_attachments(cls, messages, save_path, email_to_outlet=None, week_no=None):
        """下载附件"""
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
                
                outlet_info = None
                prefix = "UNKNOWN"
                
                if email_to_outlet:
                    outlet_info = email_to_outlet.get(sender_email)
                    
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
                
                for att in attachments:
                    filename = att.FileName
                    
                    if week_no:
                        new_filename = f"{prefix}_weekly_order_week{week_no}_{filename}"
                    else:
                        new_filename = f"{prefix}_{filename}"
                    
                    full_path = os.path.join(save_path, new_filename)
                    att.SaveAsFile(full_path)
                    result['downloaded'] += 1
            
            except Exception as e:
                print(f"Error processing attachment: {e}")
                result['skipped'] += 1
        
        return result

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
                
                total = sum(qty * price for qty, price in zip(items, prices) if price is not None)
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
    def generate_supplier_files(cls, master_file, output_folder, outlets, orders, templates, amounts, log_callback=None):
        """生成供应商文件"""
        try:
            now = datetime.now()
            next_month = now.month + 1 if now.month < 12 else 1
            year = now.year if now.month < 12 else now.year + 1
            supplier_files = []
            
            for supplier, template_ws in templates.items():
                wb = Workbook()
                wb.remove(wb.active)
                
                outlet_count = 0
                for outlet in outlets:
                    short_name = outlet["short_name"]
                    outlet_orders = orders.get(short_name, {})
                    
                    if not outlet_orders:
                        continue
                    
                    supplier_key = supplier.lower()
                    if supplier_key == "freshening":
                        order_items = outlet_orders.get("freshening", [])
                    elif supplier_key == "legacy":
                        order_items = outlet_orders.get("legacy", [])
                    else:
                        order_items = outlet_orders.get("unikleen", [])
                    
                    order_amount = amounts.get(short_name, {}).get(supplier_key, 0)
                    
                    if any(qty > 0 for qty in order_items):
                        new_ws = wb.create_sheet(title=short_name)
                        
                        for row in template_ws.iter_rows():
                            for cell in row:
                                new_cell = new_ws.cell(
                                    row=cell.row, 
                                    column=cell.column, 
                                    value=cell.value
                                )
                                
                                if cell.has_style:
                                    new_cell.font = copy.copy(cell.font)
                                    new_cell.border = copy.copy(cell.border)
                                    new_cell.fill = copy.copy(cell.fill)
                                    new_cell.number_format = cell.number_format
                                    new_cell.protection = copy.copy(cell.protection)
                                    new_cell.alignment = copy.copy(cell.alignment)
                        
                        for col_idx in range(1, template_ws.max_column + 1):
                            col_letter = get_column_letter(col_idx)
                            new_ws.column_dimensions[col_letter].width = template_ws.column_dimensions[col_letter].width
                        
                        for row_idx in range(1, template_ws.max_row + 1):
                            new_ws.row_dimensions[row_idx].height = template_ws.row_dimensions[row_idx].height
                        
                        for merged_range in template_ws.merged_cells.ranges:
                            new_ws.merge_cells(str(merged_range))
                        
                        new_ws['F5'] = outlet["full_name"]
                        new_ws['F6'] = outlet["address"]
                        
                        if supplier == "Freshening":
                            new_ws['G5'] = outlet["delivery_day"]
                        
                        if supplier == "Freshening":
                            for i, qty in enumerate(order_items[:34]):
                                new_ws[f'L{i+4}'] = qty
                            
                            new_ws['F46'] = order_amount
                            new_ws['F46'].number_format = '"$"#,##0.00'
                            
                        elif supplier == "Legacy":
                            for i, qty in enumerate(order_items[:3]):
                                new_ws[f'L{i+41}'] = qty
                            
                            cartons = sum(order_items[:3])
                            new_ws['F16'] = cartons
                            
                        else:
                            for i, qty in enumerate(order_items[:18]):
                                new_ws[f'L{i+47}'] = qty
                            
                            new_ws['F31'] = order_amount
                            new_ws['F31'].number_format = '"$"#,##0.00'
                        
                        outlet_count += 1
                
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

# ========== 入口点 ==========
if __name__ == '__main__':
    try:
        app = SushiExpressApp()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error", f"Startup failed: {e}")
        sys.exit(1)
