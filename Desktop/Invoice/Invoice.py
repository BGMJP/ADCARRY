import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser # colorchooser 임포트
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from datetime import datetime, timedelta
import os
import io
import json
from tkcalendar import DateEntry # tkcalendar 임포트

# PDF 미리보기를 위한 라이브러리
try:
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk
    PREVIEW_ENABLED = True
except ImportError:
    PREVIEW_ENABLED = False

# --- 고정 정보 ---
SUPPLIER_INFO = {
    "ko": {
        "name": "주식회사 애드캐리", "reg_num": "582-88-01950",
        "address": "서울특별시 금천구 가산디지털1로 204, 904호", "phone": "02-6925-0147", "email": "adc@adcarry.co.kr"
    },
    "en": {
        "name": "ADCARRY Corp.", "reg_num": "582-88-01950",
        "address": ["904, 204, Gasan digital 1-ro,", "Geumcheon-gu, Seoul, Republic of Korea"], "phone": "+82-2-6925-0147", "email": "adc@adcarry.co.kr"
    }
}

class InvoiceGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Gemini Invoice/Quote Generator")
        self.root.geometry("1200x800")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing) # 창 닫기 이벤트 핸들러

        # --- 변수 초기화 ---
        self.language = tk.StringVar(value="ko")
        self.doc_type = tk.StringVar(value="invoice")
        self.doc_number = tk.StringVar(value=datetime.now().strftime('%Y%m%d-%H%M%S'))
        self.doc_date = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d')) # 작성일 변수
        self.use_default_save_path = tk.BooleanVar(value=False) # 기본 저장 경로 사용 여부
        self.filename_var = tk.StringVar(value="Invoice") # 파일 이름 입력 변수
        self.add_date_prefix = tk.BooleanVar(value=True) # 날짜 접두사 추가 여부
        self.after_id = None # For debouncing preview updates
        self.preview_images = [] # 미리보기 이미지 리스트 초기화

        # --- 색상 변수 초기화 ---
        self.primary_color = tk.StringVar(value="#4A90E2") # 기본 파란색
        self.secondary_color = tk.StringVar(value="#F0F0F0") # 밝은 회색
        self.text_color = tk.StringVar(value="#333333") # 어두운 회색
        self.light_text_color = tk.StringVar(value="#FFFFFF") # 흰색
        self.border_color = tk.StringVar(value="#CCCCCC") # 중간 회색
        self.invoice_title_color = tk.StringVar(value="#FFFFFF") # 청구서 제목 글자 색상 (기본 흰색)

        # --- 다국어 텍스트 ---
        self.i18n = {
            "ko": {
                "doc_type_label": "문서 종류", "invoice": "청구서", "quote": "견적서",
                "supplier_info": "공급자 정보", "customer_info": "고객 정보", "item_info": "품목 정보",
                "company_name": "상호:", "reg_num": "사업자등록번호:", "address": "주소:", "phone": "연락처:", "email": "이메일:",
                "add_item": "품목 추가", "item": "품목", "quantity": "수량", "unit_price": "단가", "amount": "금액",
                "save_path": "저장 경로", "browse": "찾아보기...", "preview": "미리보기 생성",
                "save_pdf": "PDF로 저장", "page": "페이지", "of": "의", "preview_area": "미리보기",
                "from": "From.", "to": "To.", "doc_num": "문서번호:", "date": "작성일자:",
                "supply_amount": "공급가액", "vat": "세액 (10%)", "total_amount": "합계 금액",
                "currency_symbol": "원", "dear": "귀하",
                "issuer_info": "발행자 정보", "issuer_name": "이름:", "issuer_title": "직책:", "issuer_email": "이메일:", "issuer_phone": "연락처:",
                "item_preset": "품목 프리셋", "preset_name": "프리셋 이름:", "save_preset": "프리셋 저장", "load_preset": "프리셋 불러오기", "delete_preset": "프리셋 삭제",
                "issuer_preset": "발행자 프리셋", "save_issuer_preset": "발행자 프리셋 저장", "load_issuer_preset": "발행자 프리셋 불러오기", "delete_issuer_preset": "발행자 프리셋 삭제",
                "bank_info": "입금안내", "bank_name": "은행명:", "account_num": "계좌번호:", "account_holder": "예금주:",
                "error_fill_all": "모든 필수 필드를 채워주세요.", "error_numeric": "수량과 단가는 숫자로 입력해야 합니다.",
                "error_path": "저장 경로를 지정해주세요.", "error_preview": "먼저 미리보기를 생성해주세요.",
                "success_save": "가 다음 경로에 저장되었습니다:", "error_save": "PDF 파일을 저장하는 중 오류가 발생했습니다:\n",
                "preset_saved": "프리셋이 저장되었습니다.", "preset_loaded": "프리셋이 불러와졌습니다.", "preset_deleted": "프리셋이 삭제되었습니다.",
                "preset_name_empty": "프리셋 이름을 입력해주세요.", "preset_not_found": "선택된 프리셋을 찾을 수 없습니다.",
                "confirm_delete_preset": "선택된 프리셋을 삭제하시겠습니까?",
                "due_date": "납부기한:",
                "validity_period": "유효기한:",
                "color_settings_label": "색상 설정",
                "invoice_title_color_label": "제목 글자 색상",
                "single_item_preset": "단일 품목 프리셋",
                "save_single_item_preset": "단일 품목 프리셋 저장",
                "add_single_item": "단일 품목 추가",
                "delete_single_item_preset": "단일 품목 프리셋 삭제",
                "customer_preset": "고객 프리셋",
                "save_customer_preset": "고객 프리셋 저장",
                "load_customer_preset": "고객 프리셋 불러오기",
                "delete_customer_preset": "고객 프리셋 삭제"
            },
            "en": {
                "doc_type_label": "Document Type", "invoice": "INVOICE", "quote": "QUOTATION",
                "supplier_info": "Supplier Information", "customer_info": "Customer Information", "item_info": "Item Information",
                "company_name": "Company Name:", "reg_num": "Business Reg. No.:", "address": "Address:", "phone": "Phone:", "email": "Email:",
                "add_item": "Add Item", "item": "Item", "quantity": "Quantity", "unit_price": "Unit Price", "amount": "Amount",
                "save_path": "Save Path", "browse": "Browse...", "preview": "Generate Preview",
                "save_pdf": "Save as PDF", "page": "Page", "of": "of", "preview_area": "Preview",
                "from": "From.", "to": "To.", "doc_num": "Doc No.:", "date": "Date:",
                "supply_amount": "Subtotal", "vat": "VAT (10%)", "total_amount": "Total Amount",
                "currency_symbol": "$", "dear": "",
                "issuer_info": "Issuer Information", "issuer_name": "Name:", "issuer_title": "Title:", "issuer_email": "Email:", "issuer_phone": "Phone:",
                "item_preset": "Item Presets", "preset_name": "Preset Name:", "save_preset": "Save Preset", "load_preset": "Load Preset", "delete_preset": "Delete Preset",
                "issuer_preset": "Issuer Presets", "save_issuer_preset": "Save Issuer Preset", "load_issuer_preset": "Load Issuer Preset", "delete_issuer_preset": "Delete Issuer Preset",
                "bank_info": "Payment Information", "bank_name": "Bank Name:", "account_num": "Account No.:", "account_holder": "Account Holder:",
                "error_fill_all": "Please fill in all required fields.", "error_numeric": "Quantity and Unit Price must be numbers.",
                "error_path": "Please specify a save path.", "error_preview": "Please generate a preview first.",
                "success_save": "has been saved to:", "error_save": "An error occurred while saving the PDF:\n",
                "preset_saved": "Preset saved.", "preset_loaded": "Preset loaded.", "preset_deleted": "Preset deleted.",
                "preset_name_empty": "Please enter a preset name.", "preset_not_found": "Selected preset not found.",
                "confirm_delete_preset": "Are you sure you want to delete the selected preset?",
                "due_date": "Due Date:",
                "validity_period": "Valid Until:",
                "color_settings_label": "Color Settings",
                "invoice_title_color_label": "Title Color",
                "single_item_preset": "Single Item Preset",
                "save_single_item_preset": "Save Single Item Preset",
                "add_single_item": "Add Single Item",
                "delete_single_item_preset": "Delete Single Item Preset",
                "customer_preset": "Customer Preset",
                "save_customer_preset": "Save Customer Preset",
                "load_customer_preset": "Load Customer Preset",
                "delete_customer_preset": "Delete Customer Preset"
            }
        }

        self._register_fonts()
        self._create_widgets()
        self.language.trace_add("write", self.update_language) # Moved here
        self.doc_type.trace_add("write", self.update_language) # Moved here
        self.customer_presets = {} # 고객 프리셋 초기화
        self.single_item_presets = {} # 단일 품목 프리셋 초기화
        self.color_presets = {} # 색상 프리셋 초기화
        self.load_presets() # 모든 위젯 생성 후 프리셋 로드
        self.load_issuer_presets() # 발행자 프리셋 로드
        self.load_customer_presets() # 
        self.load_color_presets() # 색상 프리셋 로드
        self.load_window_geometry()
        self.update_language() # 초기 UI 텍스트 설정 및 미리보기 생성

    def update_language(self, *args):
        lang = self.language.get()
        t = self.i18n[lang]

        # 문서 종류 라디오 버튼 업데이트
        self.doc_type_frame.config(text=t["doc_type_label"])
        self.invoice_radio.config(text=t["invoice"])
        self.quote_radio.config(text=t["quote"])

        # 고객 정보 레이블 업데이트
        self.customer_frame.config(text=t["customer_info"])
        self.customer_labels["name"].config(text=t["company_name"])
        self.customer_labels["reg_num"].config(text=t["reg_num"])
        self.customer_labels["address"].config(text=t["address"])
        self.customer_labels["phone"].config(text=t["phone"])
        self.customer_labels["email"].config(text=t["email"])

        # 고객 프리셋 버튼 업데이트
        self.customer_preset_frame.config(text=t["customer_preset"])
        self.save_customer_preset_button.config(text=t["save_customer_preset"])
        self.load_customer_preset_button.config(text=t["load_customer_preset"])
        self.delete_customer_preset_button.config(text=t["delete_customer_preset"])

        # 발행자 정보 레이블 업데이트
        self.issuer_preset_combined_frame.config(text=t["issuer_info"])
        self.issuer_labels["name"].config(text=t["issuer_name"])
        self.issuer_labels["title"].config(text=t["issuer_title"])
        self.issuer_labels["email"].config(text=t["issuer_email"])
        self.issuer_labels["phone"].config(text=t["issuer_phone"])

        # 품목 정보 레이블 업데이트
        self.item_preset_combined_frame.config(text=t["item_info"])
        self.add_item_button.config(text=t["add_item"])
        self.update_item_labels()

        # 저장 경로 버튼 업데이트
        self.path_frame.config(text=t["save_path"])
        self.browse_button.config(text=t["browse"])

        # 미리보기 및 저장 버튼 업데이트
        self.preview_button.config(text=t["preview"])
        self.save_pdf_button.config(text=t["save_pdf"])
        self.page_label.config(text=f"{t['page']} {self.current_page + 1 if hasattr(self, 'current_page') else 1} {t['of']} {len(self.preview_images) if hasattr(self, 'preview_images') else 1}")

        # 미리보기 영역 프레임 업데이트
        self.preview_label_frame.config(text=t["preview_area"])

        # 품목 프리셋 버튼 업데이트
        self.preset_frame.config(text=t["item_preset"])
        self.save_preset_button.config(text=t["save_preset"])
        self.load_preset_button.config(text=t["load_preset"])
        self.delete_preset_button.config(text=t["delete_preset"])

        # 단일 품목 프리셋 버튼 업데이트
        self.single_item_preset_frame.config(text=t["single_item_preset"])
        self.save_single_item_preset_button.config(text=t["save_single_item_preset"])
        self.add_single_item_button.config(text=t["add_single_item"])
        self.delete_single_item_preset_button.config(text=t["delete_single_item_preset"])

        # 발행자 프리셋 버튼 업데이트
        self.issuer_preset_frame.config(text=t["issuer_preset"])
        self.save_issuer_preset_button.config(text=t["save_issuer_preset"])
        self.load_issuer_preset_button.config(text=t["load_issuer_preset"])
        self.delete_issuer_preset_button.config(text=t["delete_issuer_preset"])

        # 색상 설정 프레임 업데이트
        self.color_settings_frame.config(text=t["color_settings_label"])

        # 계좌 정보 프레임 업데이트 (주석 처리된 부분은 제외)
        self.bank_frame.config(text=t["bank_info"])

    def on_closing(self):
        self.save_window_geometry()
        self.root.destroy()

    def save_window_geometry(self):
        geom = self.root.geometry()
        with open("window_geometry.json", "w") as f:
            json.dump({"geometry": geom}, f)

    def load_window_geometry(self):
        if os.path.exists("window_geometry.json"):
            with open("window_geometry.json", "r") as f:
                try:
                    data = json.load(f)
                    self.root.geometry(data["geometry"])
                except json.JSONDecodeError:
                    pass # 파일이 손상된 경우 무시

    def _pick_color(self, color_var):
        color_code = colorchooser.askcolor(color=color_var.get())
        if color_code[1]: # 사용자가 색상을 선택하고 '확인'을 눌렀을 경우
            color_var.set(color_code[1].upper()) # 헥스 코드 (대문자로)

    def _schedule_preview_update(self):
        if self.after_id:
            self.root.after_cancel(self.after_id)
        self.after_id = self.root.after(500, self.generate_preview) # 500ms (0.5초) 지연 후 미리보기 생성

    def _register_fonts(self):
        import sys
        
        def resource_path(relative_path):
                try:
                    base_path = sys._MEIPASS
                except Exception:
                    base_path = os.path.abspath(".")
                return os.path.join(base_path, relative_path)
        font_path = resource_path(os.path.join("fonts", "malgun.ttf"))
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('MalgunGothic', font_path))
            pdfmetrics.registerFont(TTFont('MalgunGothic-Bold', font_path))
            pdfmetrics.registerFontFamily('MalgunGothic', normal='MalgunGothic', bold='MalgunGothic-Bold')
            self.default_font = 'MalgunGothic'
            self.default_font_bold = 'MalgunGothic-Bold'
        else:
            messagebox.showwarning("Font Warning", "Malgun Gothic font not found. PDF text may not display correctly.")
            self.default_font = 'Helvetica' # Fallback
            self.default_font_bold = 'Helvetica-Bold'

    def _create_entry_label(self, parent_frame, row, entry_widget, width=None):
        # 이 메서드는 레이블과 엔트리 위젯을 생성하고 그리드에 배치합니다.
        label = ttk.Label(parent_frame, text="") # 실제 텍스트는 호출하는 곳에서 설정
        label.grid(row=row, column=0, sticky=tk.W, pady=2)
        entry_widget.grid(row=row, column=1, sticky=tk.W, pady=2)
        if width:
            entry_widget.config(width=width)
        return label # 레이블 위젯을 반환하여 나중에 텍스트를 업데이트할 수 있도록 합니다.

    def _create_color_picker_row(self, parent_frame, label_text, color_var, row):
        label = ttk.Label(parent_frame, text=label_text)
        label.grid(row=row, column=0, sticky=tk.W, pady=2)

        color_entry = ttk.Entry(parent_frame, textvariable=color_var, width=10, state="readonly")
        color_entry.grid(row=row, column=1, sticky=tk.W, pady=2, padx=5)

        color_button = ttk.Button(parent_frame, text="선택", command=lambda: self._pick_color(color_var))
        color_button.grid(row=row, column=2, sticky=tk.W, pady=2)
        color_var.trace_add("write", lambda *args: self._schedule_preview_update())

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 스크롤 가능한 캔버스 생성
        canvas = tk.Canvas(main_frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 스크롤바 생성 및 캔버스에 연결
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        # left_panel을 캔버스 안에 생성
        left_panel = ttk.Frame(canvas, padding="10")
        canvas.create_window((0, 0), window=left_panel, anchor="nw")

        # left_panel의 크기가 변경될 때마다 스크롤 영역 업데이트
        left_panel.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # --- 상단 설정 (언어, 문서종류, 로고) ---
        settings_frame = ttk.Frame(left_panel)
        settings_frame.pack(fill=tk.X, pady=5)
        ttk.Radiobutton(settings_frame, text="한국어", variable=self.language, value="ko", command=self._schedule_preview_update).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(settings_frame, text="English", variable=self.language, value="en", command=self._schedule_preview_update).pack(side=tk.LEFT, padx=5)
        
        self.doc_type_frame = ttk.LabelFrame(left_panel, padding="10")
        self.doc_type_frame.pack(fill=tk.X, pady=10)
        self.invoice_radio = ttk.Radiobutton(self.doc_type_frame, variable=self.doc_type, value="invoice", command=self._schedule_preview_update)
        self.invoice_radio.pack(side=tk.LEFT, padx=5)
        self.quote_radio = ttk.Radiobutton(self.doc_type_frame, variable=self.doc_type, value="quote", command=self._schedule_preview_update)
        self.quote_radio.pack(side=tk.LEFT, padx=5)

        # --- 문서 정보 및 발행자 정보 컨테이너 ---
        doc_issuer_container_frame = ttk.Frame(left_panel)
        doc_issuer_container_frame.pack(fill=tk.X, pady=10)

        # --- 문서 번호 및 작성일 ---
        doc_info_frame = ttk.LabelFrame(doc_issuer_container_frame, text="문서 정보", padding="10")
        doc_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        ttk.Label(doc_info_frame, text="번호:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.doc_number_entry = ttk.Entry(doc_info_frame, textvariable=self.doc_number, width=30)
        self.doc_number_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.doc_number_entry.bind("<KeyRelease>", lambda e: self._schedule_preview_update())

        ttk.Label(doc_info_frame, text="작성일:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.doc_date_entry = DateEntry(doc_info_frame, selectmode='day', textvariable=self.doc_date, 
                                        date_pattern='yyyy-mm-dd', width=27, background='darkblue',
                                        foreground='white', borderwidth=2)
        self.doc_date_entry.grid(row=1, column=1, sticky=tk.W, pady=2)
        self.doc_date_entry.bind("<<DateSelected>>", lambda e: self._schedule_preview_update())

        # --- 발행자 정보 및 프리셋 ---
        self.issuer_preset_combined_frame = ttk.LabelFrame(doc_issuer_container_frame, padding="10")
        self.issuer_preset_combined_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # --- 고객 정보 및 프리셋 컨테이너 ---
        customer_container_frame = ttk.Frame(left_panel)
        customer_container_frame.pack(fill=tk.X, pady=10)

        # 고객 정보
        self.customer_frame = ttk.LabelFrame(customer_container_frame, padding="10")
        self.customer_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 고객 정보 입력 필드
        self.customer_name = ttk.Entry(self.customer_frame, width=30)
        self.customer_reg_num = ttk.Entry(self.customer_frame, width=30)
        self.customer_address = ttk.Entry(self.customer_frame, width=30)
        self.customer_phone = ttk.Entry(self.customer_frame, width=30)
        self.customer_email = ttk.Entry(self.customer_frame, width=30)

        self.customer_labels = {
            "name": self._create_entry_label(self.customer_frame, 0, self.customer_name, width=30),
            "reg_num": self._create_entry_label(self.customer_frame, 1, self.customer_reg_num, width=30),
            "address": self._create_entry_label(self.customer_frame, 2, self.customer_address, width=30),
            "phone": self._create_entry_label(self.customer_frame, 3, self.customer_phone, width=30),
            "email": self._create_entry_label(self.customer_frame, 4, self.customer_email, width=30)
        }
        self.customer_name.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.customer_reg_num.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.customer_address.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.customer_phone.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.customer_email.bind("<KeyRelease>", lambda e: self._schedule_preview_update())

        # 고객 프리셋
        self.customer_preset_frame = ttk.LabelFrame(customer_container_frame, padding="5")
        self.customer_preset_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        ttk.Label(self.customer_preset_frame, text="프리셋 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.customer_preset_name_entry = ttk.Entry(self.customer_preset_frame, width=20)
        self.customer_preset_name_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.customer_preset_combobox = ttk.Combobox(self.customer_preset_frame, state="readonly", width=25)
        self.customer_preset_combobox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.customer_preset_combobox.bind("<<ComboboxSelected>>", lambda e: (self.on_customer_preset_selected(e), self._schedule_preview_update()))
        customer_preset_button_frame = ttk.Frame(self.customer_preset_frame)
        customer_preset_button_frame.grid(row=2, column=0, columnspan=2, pady=5)
        self.save_customer_preset_button = ttk.Button(customer_preset_button_frame, text="고객 프리셋 저장", command=self.add_customer_preset)
        self.save_customer_preset_button.pack(side=tk.LEFT, padx=2)
        self.load_customer_preset_button = ttk.Button(customer_preset_button_frame, text="고객 프리셋 불러오기", command=lambda: (self.apply_customer_preset(), self._schedule_preview_update()))
        self.load_customer_preset_button.pack(side=tk.LEFT, padx=2)
        self.delete_customer_preset_button = ttk.Button(customer_preset_button_frame, text="고객 프리셋 삭제", command=self.delete_customer_preset)
        self.delete_customer_preset_button.pack(side=tk.LEFT, padx=2)

        # 발행자 정보
        self.issuer_frame = ttk.Frame(self.issuer_preset_combined_frame)
        self.issuer_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)
        self.issuer_name = ttk.Entry(self.issuer_frame, width=30)
        self.issuer_title = ttk.Entry(self.issuer_frame, width=30)
        self.issuer_email = ttk.Entry(self.issuer_frame, width=30)
        self.issuer_phone = ttk.Entry(self.issuer_frame, width=30)
        self.issuer_labels = {
            "name": self._create_entry_label(self.issuer_frame, 0, self.issuer_name, width=30),
            "title": self._create_entry_label(self.issuer_frame, 1, self.issuer_title, width=30),
            "email": self._create_entry_label(self.issuer_frame, 2, self.issuer_email, width=30),
            "phone": self._create_entry_label(self.issuer_frame, 3, self.issuer_phone, width=30)
        }
        self.issuer_name.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.issuer_title.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.issuer_email.bind("<KeyRelease>", lambda e: self._schedule_preview_update())
        self.issuer_phone.bind("<KeyRelease>", lambda e: self._schedule_preview_update())

        # 발행자 프리셋
        self.issuer_preset_frame = ttk.LabelFrame(self.issuer_preset_combined_frame, padding="5")
        self.issuer_preset_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y)
        ttk.Label(self.issuer_preset_frame, text="프리셋 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.issuer_preset_name_entry = ttk.Entry(self.issuer_preset_frame, width=20)
        self.issuer_preset_name_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.issuer_preset_combobox = ttk.Combobox(self.issuer_preset_frame, state="readonly", width=25)
        self.issuer_preset_combobox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.issuer_preset_combobox.bind("<<ComboboxSelected>>", lambda e: (self.on_issuer_preset_selected(e), self._schedule_preview_update()))
        issuer_preset_button_frame = ttk.Frame(self.issuer_preset_frame)
        issuer_preset_button_frame.grid(row=2, column=0, columnspan=2, pady=5)
        self.save_issuer_preset_button = ttk.Button(issuer_preset_button_frame, text="발행자 프리셋 저장", command=self.add_issuer_preset)
        self.save_issuer_preset_button.pack(side=tk.LEFT, padx=2)
        self.load_issuer_preset_button = ttk.Button(issuer_preset_button_frame, text="발행자 프리셋 불러오기", command=lambda: (self.apply_issuer_preset(), self._schedule_preview_update()))
        self.load_issuer_preset_button.pack(side=tk.LEFT, padx=2)
        self.delete_issuer_preset_button = ttk.Button(issuer_preset_button_frame, text="발행자 프리셋 삭제", command=self.delete_issuer_preset)
        self.delete_issuer_preset_button.pack(side=tk.LEFT, padx=2)


        # --- 품목 정보 ---
        self.items_frame = ttk.Frame(left_panel)
        self.items_frame.pack(fill=tk.X, pady=10)
        # 품목 정보 입력란 열 너비 설정
        self.items_frame.grid_columnconfigure(0, weight=1) # 품목 이름
        self.items_frame.grid_columnconfigure(1, weight=1) # 수량
        self.items_frame.grid_columnconfigure(2, weight=1) # 단가
        self.items_frame.grid_columnconfigure(3, weight=1) # 금액
        self.items = []

        # 품목 헤더 레이블 생성 (한 번만 생성)
        self.item_header_labels = {
            "item": ttk.Label(self.items_frame, text="품목"),
            "quantity": ttk.Label(self.items_frame, text="수량", anchor=tk.CENTER),
            "unit_price": ttk.Label(self.items_frame, text="단가", anchor=tk.CENTER),
            "amount": ttk.Label(self.items_frame, text="금액", anchor=tk.CENTER)
        }
        self.item_header_labels["item"].grid(row=0, column=0, padx=2, sticky=tk.W + tk.E)
        self.item_header_labels["quantity"].grid(row=0, column=1, padx=2, sticky=tk.W + tk.E)
        self.item_header_labels["unit_price"].grid(row=0, column=2, padx=2, sticky=tk.W + tk.E)
        self.item_header_labels["amount"].grid(row=0, column=3, padx=2, sticky=tk.W + tk.E)

        item_buttons_frame = ttk.Frame(self.items_frame)
        item_buttons_frame.grid(row=100, column=0, columnspan=4, pady=10)

        self.add_item_button = ttk.Button(item_buttons_frame, command=lambda: (self.add_item_row(len(self.items)), self._schedule_preview_update()))
        self.add_item_button.pack(side=tk.LEFT, padx=5)

        self.remove_item_button = ttk.Button(item_buttons_frame, text="품목 제거", command=self.remove_last_item_row)
        self.remove_item_button.pack(side=tk.LEFT, padx=5)
        self.add_item_row(0) # Initial item row

        # --- 품목 프리셋 컨테이너 (다중 품목 및 단일 품목 프리셋을 포함) ---
        self.item_preset_combined_frame = ttk.LabelFrame(left_panel, text="품목 정보", padding="10")
        self.item_preset_combined_frame.pack(fill=tk.X, pady=10)

        # 품목 프리셋
        self.preset_frame = ttk.LabelFrame(self.item_preset_combined_frame, padding="5")
        self.preset_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y, expand=True)
        ttk.Label(self.preset_frame, text="프리셋 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.preset_name_entry = ttk.Entry(self.preset_frame, width=20)
        self.preset_name_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.preset_combobox = ttk.Combobox(self.preset_frame, state="readonly", width=25)
        self.preset_combobox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.preset_combobox.bind("<<ComboboxSelected>>", lambda e: (self.on_preset_selected(e), self._schedule_preview_update()))
        preset_button_frame = ttk.Frame(self.preset_frame)
        preset_button_frame.grid(row=2, column=0, columnspan=2, pady=5)
        self.save_preset_button = ttk.Button(preset_button_frame, text="프리셋 저장", command=self.add_preset)
        self.save_preset_button.pack(side=tk.LEFT, padx=2)
        self.load_preset_button = ttk.Button(preset_button_frame, text="프리셋 불러오기", command=lambda: (self.apply_preset(), self._schedule_preview_update()))
        self.load_preset_button.pack(side=tk.LEFT, padx=2)
        self.delete_preset_button = ttk.Button(preset_button_frame, text="프리셋 삭제", command=self.delete_preset)
        self.delete_preset_button.pack(side=tk.LEFT, padx=2)

        # 단일 품목 프리셋
        self.single_item_preset_frame = ttk.LabelFrame(self.item_preset_combined_frame, text="단일 품목 프리셋", padding="5")
        self.single_item_preset_frame.pack(side=tk.LEFT, padx=10, fill=tk.Y, expand=True)

        ttk.Label(self.single_item_preset_frame, text="프리셋 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.single_preset_name_entry = ttk.Entry(self.single_item_preset_frame, width=20)
        self.single_preset_name_entry.grid(row=0, column=1, sticky=tk.W, pady=2)

        ttk.Label(self.single_item_preset_frame, text="품목명:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.single_preset_item_name_entry = ttk.Entry(self.single_item_preset_frame, width=20)
        self.single_preset_item_name_entry.grid(row=1, column=1, sticky=tk.W, pady=2)

        ttk.Label(self.single_item_preset_frame, text="수량:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.single_preset_item_quantity_entry = ttk.Entry(self.single_item_preset_frame, width=20)
        self.single_preset_item_quantity_entry.grid(row=2, column=1, sticky=tk.W, pady=2)

        ttk.Label(self.single_item_preset_frame, text="단가:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.single_preset_item_unit_price_entry = ttk.Entry(self.single_item_preset_frame, width=20)
        self.single_preset_item_unit_price_entry.grid(row=3, column=1, sticky=tk.W, pady=2)

        self.single_preset_combobox = ttk.Combobox(self.single_item_preset_frame, state="readonly", width=25)
        self.single_preset_combobox.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.single_preset_combobox.bind("<<ComboboxSelected>>", lambda e: (self.on_single_item_preset_selected(e), self._schedule_preview_update()))

        single_preset_button_frame = ttk.Frame(self.single_item_preset_frame)
        single_preset_button_frame.grid(row=5, column=0, columnspan=2, pady=5)
        self.save_single_item_preset_button = ttk.Button(single_preset_button_frame, text="단일 품목 프리셋 저장", command=self.add_single_item_preset)
        self.save_single_item_preset_button.pack(side=tk.LEFT, padx=2)
        self.add_single_item_button = ttk.Button(single_preset_button_frame, text="단일 품목 추가", command=lambda: (self.apply_single_item_preset(), self._schedule_preview_update()))
        self.add_single_item_button.pack(side=tk.LEFT, padx=2)
        self.delete_single_item_preset_button = ttk.Button(single_preset_button_frame, text="단일 품목 프리셋 삭제", command=self.delete_single_item_preset)
        self.delete_single_item_preset_button.pack(side=tk.LEFT, padx=2)

        # --- 계좌 정보 ---
        self.bank_frame = ttk.LabelFrame(left_panel, padding="10")
        self.bank_frame.pack(fill=tk.X, pady=10)
        # 계좌 정보는 PDF에 직접 하드코딩되므로 GUI 입력 필드는 제거합니다。
        # self.bank_name = ttk.Entry(self.bank_frame, width=30)
        # self.account_num = ttk.Entry(self.bank_frame, width=30)
        # self.account_holder = ttk.Entry(self.bank_frame, width=30)
        # self.bank_labels = {
        #     "name": self._create_entry_label(self.bank_frame, 0, self.bank_name, width=30),
        #     "account_num": self._create_entry_label(self.bank_frame, 1, self.account_num, width=30),
        #     "account_holder": self._create_entry_label(self.bank_frame, 2, self.account_holder, width=30)
        # }

        # --- 설정 컨테이너 (색상, 저장 경로, 파일 이름) ---
        settings_container_frame = ttk.Frame(left_panel, padding="10")
        settings_container_frame.pack(fill=tk.X, pady=10)

        # 색상 설정
        self.color_settings_frame = ttk.LabelFrame(settings_container_frame, text=self.i18n[self.language.get()]["color_settings_label"], padding="10")
        self.color_settings_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        self._create_color_picker_row(self.color_settings_frame, "기본 색상:", self.primary_color, 0)
        self._create_color_picker_row(self.color_settings_frame, "보조 색상:", self.secondary_color, 1)
        self._create_color_picker_row(self.color_settings_frame, "텍스트 색상:", self.text_color, 2)
        self._create_color_picker_row(self.color_settings_frame, "밝은 텍스트 색상:", self.light_text_color, 3)
        self._create_color_picker_row(self.color_settings_frame, "테두리 색상:", self.border_color, 4)
        self._create_color_picker_row(self.color_settings_frame, self.i18n[self.language.get()]["invoice_title_color_label"], self.invoice_title_color, 5)

        # 색상 프리셋
        color_preset_frame = ttk.Frame(self.color_settings_frame)
        color_preset_frame.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=5)
        ttk.Label(color_preset_frame, text="프리셋 이름:").pack(side=tk.LEFT, padx=2)
        self.color_preset_name_entry = ttk.Entry(color_preset_frame, width=15)
        self.color_preset_name_entry.pack(side=tk.LEFT, padx=2)
        self.color_preset_combobox = ttk.Combobox(color_preset_frame, state="readonly", width=15)
        self.color_preset_combobox.pack(side=tk.LEFT, padx=2)
        self.color_preset_combobox.bind("<<ComboboxSelected>>", lambda e: (self.on_color_preset_selected(e), self._schedule_preview_update()))

        color_preset_button_frame = ttk.Frame(self.color_settings_frame)
        color_preset_button_frame.grid(row=6, column=0, columnspan=3, pady=5)
        ttk.Button(color_preset_button_frame, text="프리셋 저장", command=self.add_color_preset).pack(side=tk.LEFT, padx=2)
        ttk.Button(color_preset_button_frame, text="프리셋 불러오기", command=lambda: (self.apply_color_preset(), self._schedule_preview_update())).pack(side=tk.LEFT, padx=2)
        ttk.Button(color_preset_button_frame, text="프리셋 삭제", command=self.delete_color_preset).pack(side=tk.LEFT, padx=2)
        
        # 저장 경로 및 파일 이름 설정 컨테이너
        path_filename_container_frame = ttk.Frame(settings_container_frame)
        path_filename_container_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        # 저장 경로
        self.path_frame = ttk.LabelFrame(path_filename_container_frame, padding="10")
        self.path_frame.pack(fill=tk.X, pady=5)
        self.save_path = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop"))
        path_entry = ttk.Entry(self.path_frame, textvariable=self.save_path, width=40, state="readonly", name="path_entry") # Added name for easier access
        path_entry.grid(row=0, column=0, sticky="we", padx=5)
        self.browse_button = ttk.Button(self.path_frame, command=self.select_save_path)
        self.browse_button.grid(row=0, column=1, padx=5)

        self.default_path_checkbox = ttk.Checkbutton(self.path_frame, text="기본 저장 경로로 사용", variable=self.use_default_save_path, command=self.toggle_save_path_state)
        self.default_path_checkbox.grid(row=1, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)
        self.load_default_save_path_setting() # 설정 로드

        # 파일 이름 설정
        filename_frame = ttk.LabelFrame(path_filename_container_frame, text="파일 이름 설정", padding="10")
        filename_frame.pack(fill=tk.X, pady=5)
        ttk.Label(filename_frame, text="파일 이름:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.filename_entry = ttk.Entry(filename_frame, textvariable=self.filename_var, width=40)
        self.filename_entry.grid(row=0, column=1, sticky=tk.W, pady=2)
        self.add_date_prefix_checkbox = ttk.Checkbutton(filename_frame, text="날짜 접두사 추가 (YYYYMMDD_)", variable=self.add_date_prefix)
        self.add_date_prefix_checkbox.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=2)

        # --- 미리보기 및 저장 버튼 ---
        action_frame = ttk.Frame(left_panel, padding="10")
        action_frame.pack(fill=tk.X, pady=10)

        self.preview_button = ttk.Button(action_frame, command=self.generate_preview)
        self.preview_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.save_pdf_button = ttk.Button(action_frame, command=self.save_pdf_from_preview)
        self.save_pdf_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        # --- 미리보기 영역 ---
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10)

        preview_frame = ttk.LabelFrame(right_panel, text="미리보기", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        self.preview_label_frame = preview_frame # 직접 참조 저장

        self.preview_canvas = tk.Canvas(preview_frame, bg="white", bd=2, relief="groove")
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_canvas.bind("<Configure>", lambda e: self.update_preview_display())

        # 페이지 탐색 버튼
        page_nav_frame = ttk.Frame(preview_frame)
        page_nav_frame.pack(pady=5)

        self.prev_button = ttk.Button(page_nav_frame, text="< 이전", command=self.show_previous_page, state=tk.DISABLED)
        self.prev_button.pack(side=tk.LEFT, padx=5)

        self.page_label = ttk.Label(page_nav_frame, text="페이지 1/1")
        self.page_label.pack(side=tk.LEFT, padx=5)

        self.next_button = ttk.Button(page_nav_frame, text="다음 >", command=self.show_next_page, state=tk.DISABLED)
        self.next_button.pack(side=tk.LEFT, padx=5)

    def toggle_save_path_state(self):
        if self.use_default_save_path.get():
            # 체크박스 선택 시: 현재 save_path를 last_save_path로 저장하고, 경로 입력 필드를 비활성화
            self.path_frame.children["path_entry"].config(state="readonly")
            self.browse_button.config(state=tk.DISABLED)
            self.save_default_save_path_setting()
        else:
            # 체크박스 해제 시: 경로 입력 필드를 활성화
            self.path_frame.children["path_entry"].config(state="enabled")
            self.browse_button.config(state=tk.NORMAL)
            self.save_default_save_path_setting()

    def save_default_save_path_setting(self):
        settings_file = "settings.json"
        settings = {}
        if os.path.exists(settings_file):
            try:
                with open(settings_file, "r", encoding="utf-8") as f:
                    settings = json.load(f)
            except json.JSONDecodeError:
                pass
        settings["use_default_save_path"] = self.use_default_save_path.get()
        settings["last_save_path"] = self.save_path.get() # 마지막 저장 경로 추가
        with open(settings_file, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)

    def load_default_save_path_setting(self):
        settings_file = "settings.json"
        if os.path.exists(settings_file):
            try:
                with open(settings_file, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                    if "use_default_save_path" in settings:
                        self.use_default_save_path.set(settings["use_default_save_path"])
                    if "last_save_path" in settings and settings["use_default_save_path"]:
                        self.save_path.set(settings["last_save_path"])
                    self.toggle_save_path_state() # 상태에 따라 위젯 업데이트
            except json.JSONDecodeError:
                pass

    def select_save_path(self):
        path = filedialog.askdirectory()
        if path:
            self.save_path.set(path)

    def load_presets(self):
        preset_file = "item_presets.json"
        if os.path.exists(preset_file):
            try:
                with open(preset_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.item_presets = data.get("multi_item_presets", {})
                    self.single_item_presets = data.get("single_item_presets", {})
            except json.JSONDecodeError:
                self.item_presets = {}
                self.single_item_presets = {}
        self.update_preset_combobox()
        self.update_single_item_preset_combobox()

    def save_presets(self):
        preset_file = "item_presets.json"
        with open(preset_file, "w", encoding="utf-8") as f:
            json.dump({"multi_item_presets": self.item_presets, "single_item_presets": self.single_item_presets}, f, indent=4, ensure_ascii=False)

    def update_preset_combobox(self):
        self.preset_combobox["values"] = list(self.item_presets.keys())
        if self.item_presets:
            self.preset_combobox.set(list(self.item_presets.keys())[0])
        else:
            self.preset_combobox.set("")

    def add_preset(self):
        preset_name = self.preset_name_entry.get().strip()
        if not preset_name:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_name_empty"])
            return

        current_items = []
        for item_widget in self.items:
            name = item_widget['name'].get()
            quantity = item_widget['quantity'].get()
            unit_price = item_widget['unit_price'].get()
            if name or quantity or unit_price: # 비어있지 않은 품목만 저장
                current_items.append({
                    "name": name,
                    "quantity": quantity,
                    "unit_price": unit_price
                })
        
        if not current_items:
            messagebox.showwarning("Warning", "저장할 품목이 없습니다.")
            return

        self.item_presets[preset_name] = current_items
        self.save_presets()
        self.update_preset_combobox()
        messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_saved"])

    def apply_preset(self):
        preset_name = self.preset_combobox.get()
        if not preset_name or preset_name not in self.item_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        # 기존 품목 필드 초기화
        for item_widget in self.items:
            item_widget['name'].destroy()
            item_widget['quantity'].destroy()
            item_widget['unit_price'].destroy()
            item_widget['amount'].destroy()
        self.items = []

        # 프리셋 데이터로 품목 필드 채우기
        for i, item_data in enumerate(self.item_presets[preset_name]):
            self.add_item_row(i)
            self.items[i]['name'].insert(0, item_data['name'])
            self.items[i]['quantity'].insert(0, item_data['quantity'])
            self.items[i]['unit_price'].insert(0, item_data['unit_price'])
            # 금액 자동 업데이트 트리거
            self.items[i]['quantity'].event_generate("<KeyRelease>")

        # messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_loaded"])

    def delete_preset(self):
        preset_name = self.preset_combobox.get()
        if not preset_name or preset_name not in self.item_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        if messagebox.askyesno("Confirm Delete", self.i18n[self.language.get()]["confirm_delete_preset"]):
            del self.item_presets[preset_name]
            self.save_presets()
            self.update_preset_combobox()
            messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_deleted"])

    def on_preset_selected(self, event):
        # 콤보박스에서 프리셋 선택 시, 프리셋 이름 엔트리에 자동 입력
        selected_preset = self.preset_combobox.get()
        self.preset_name_entry.delete(0, tk.END)
        self.preset_name_entry.insert(0, selected_preset)

    # --- 단일 품목 프리셋 관련 함수 ---
    def update_single_item_preset_combobox(self):
        self.single_preset_combobox["values"] = list(self.single_item_presets.keys())
        if self.single_item_presets:
            self.single_preset_combobox.set(list(self.single_item_presets.keys())[0])
        else:
            self.single_preset_combobox.set("")

    def add_single_item_preset(self):
        preset_name = self.single_preset_name_entry.get().strip()
        if not preset_name:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_name_empty"])
            return

        item_name = self.single_preset_item_name_entry.get().strip()
        item_quantity = self.single_preset_item_quantity_entry.get().strip()
        item_unit_price = self.single_preset_item_unit_price_entry.get().strip()

        if not item_name and not item_quantity and not item_unit_price:
            messagebox.showwarning("Warning", "저장할 품목 정보를 입력해주세요.")
            return

        self.single_item_presets[preset_name] = {
            "name": item_name,
            "quantity": item_quantity,
            "unit_price": item_unit_price
        }
        self.save_presets()
        self.update_single_item_preset_combobox()
        messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_saved"])

    def apply_single_item_preset(self):
        preset_name = self.single_preset_combobox.get()
        if not preset_name or preset_name not in self.single_item_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        item_data = self.single_item_presets[preset_name]
        
        # 새 품목 행 추가
        new_item_index = len(self.items)
        self.add_item_row(new_item_index)

        # 새로 추가된 품목 행에 프리셋 데이터 채우기
        self.items[new_item_index]['name'].insert(0, item_data['name'])
        self.items[new_item_index]['quantity'].insert(0, item_data['quantity'])
        self.items[new_item_index]['unit_price'].insert(0, item_data['unit_price'])
        
        # 금액 자동 업데이트 트리거
        self.items[new_item_index]['quantity'].event_generate("<KeyRelease>")

    def delete_single_item_preset(self):
        preset_name = self.single_preset_combobox.get()
        if not preset_name or preset_name not in self.single_item_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        if messagebox.askyesno("Confirm Delete", self.i18n[self.language.get()]["confirm_delete_preset"]):
            del self.single_item_presets[preset_name]
            self.save_presets()
            self.update_single_item_preset_combobox()
            messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_deleted"])

    def on_single_item_preset_selected(self, event):
        selected_preset = self.single_preset_combobox.get()
        self.single_preset_name_entry.delete(0, tk.END)
        self.single_preset_name_entry.insert(0, selected_preset)

        if selected_preset in self.single_item_presets:
            item_data = self.single_item_presets[selected_preset]
            self.single_preset_item_name_entry.delete(0, tk.END)
            self.single_preset_item_name_entry.insert(0, item_data.get('name', ''))
            self.single_preset_item_quantity_entry.delete(0, tk.END)
            self.single_preset_item_quantity_entry.insert(0, item_data.get('quantity', ''))
            self.single_preset_item_unit_price_entry.delete(0, tk.END)
            self.single_preset_item_unit_price_entry.insert(0, item_data.get('unit_price', ''))

    def load_issuer_presets(self):
        preset_file = "issuer_presets.json"
        if os.path.exists(preset_file):
            try:
                with open(preset_file, "r", encoding="utf-8") as f:
                    self.issuer_presets = json.load(f)
            except json.JSONDecodeError:
                self.issuer_presets = {}
        self.update_issuer_preset_combobox()

    def save_issuer_presets(self):
        preset_file = "issuer_presets.json"
        with open(preset_file, "w", encoding="utf-8") as f:
            json.dump(self.issuer_presets, f, indent=4, ensure_ascii=False)

    def update_issuer_preset_combobox(self):
        self.issuer_preset_combobox["values"] = list(self.issuer_presets.keys())
        if self.issuer_presets:
            self.issuer_preset_combobox.set(list(self.issuer_presets.keys())[0])
        else:
            self.issuer_preset_combobox.set("")

    def add_issuer_preset(self):
        preset_name = self.issuer_preset_name_entry.get().strip()
        if not preset_name:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_name_empty"])
            return

        current_issuer_info = {
            "name": self.issuer_name.get(),
            "title": self.issuer_title.get(),
            "email": self.issuer_email.get(),
            "phone": self.issuer_phone.get()
        }
        
        if not any(current_issuer_info.values()):
            messagebox.showwarning("Warning", "저장할 발행자 정보가 없습니다.")
            return

        self.issuer_presets[preset_name] = current_issuer_info
        self.save_issuer_presets()
        self.update_issuer_preset_combobox()
        messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_saved"])

    def apply_issuer_preset(self):
        preset_name = self.issuer_preset_combobox.get()
        if not preset_name or preset_name not in self.issuer_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        issuer_data = self.issuer_presets[preset_name]
        self.issuer_name.delete(0, tk.END); self.issuer_name.insert(0, issuer_data['name'])
        self.issuer_title.delete(0, tk.END); self.issuer_title.insert(0, issuer_data['title'])
        self.issuer_email.delete(0, tk.END); self.issuer_email.insert(0, issuer_data['email'])
        self.issuer_phone.delete(0, tk.END); self.issuer_phone.insert(0, issuer_data['phone'])

        # messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_loaded"])

    def delete_issuer_preset(self):
        preset_name = self.issuer_preset_combobox.get()
        if not preset_name or preset_name not in self.issuer_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        if messagebox.askyesno("Confirm Delete", self.i18n[self.language.get()]["confirm_delete_preset"]):
            del self.issuer_presets[preset_name]
            self.save_issuer_presets()
            self.update_issuer_preset_combobox()
            messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_deleted"])

    def on_issuer_preset_selected(self, event):
        selected_preset = self.issuer_preset_combobox.get()
        self.issuer_preset_name_entry.delete(0, tk.END)
        self.issuer_preset_name_entry.insert(0, selected_preset)

    # --- 고객 프리셋 관련 함수 ---
    def load_customer_presets(self):
        preset_file = "customer_presets.json"
        if os.path.exists(preset_file):
            try:
                with open(preset_file, "r", encoding="utf-8") as f:
                    self.customer_presets = json.load(f)
            except json.JSONDecodeError:
                self.customer_presets = {}
        self.update_customer_preset_combobox()

    def save_customer_presets(self):
        preset_file = "customer_presets.json"
        with open(preset_file, "w", encoding="utf-8") as f:
            json.dump(self.customer_presets, f, indent=4, ensure_ascii=False)

    def update_customer_preset_combobox(self):
        self.customer_preset_combobox["values"] = list(self.customer_presets.keys())
        if self.customer_presets:
            self.customer_preset_combobox.set(list(self.customer_presets.keys())[0])
        else:
            self.customer_preset_combobox.set("")

    def add_customer_preset(self):
        preset_name = self.customer_preset_name_entry.get().strip()
        if not preset_name:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_name_empty"])
            return

        current_customer_info = {
            "name": self.customer_name.get(),
            "reg_num": self.customer_reg_num.get(),
            "address": self.customer_address.get(),
            "phone": self.customer_phone.get(),
            "email": self.customer_email.get()
        }
        
        if not any(current_customer_info.values()):
            messagebox.showwarning("Warning", "저장할 고객 정보가 없습니다.")
            return

        self.customer_presets[preset_name] = current_customer_info
        self.save_customer_presets()
        self.update_customer_preset_combobox()
        messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_saved"])

    def apply_customer_preset(self):
        preset_name = self.customer_preset_combobox.get()
        if not preset_name or preset_name not in self.customer_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        customer_data = self.customer_presets[preset_name]
        self.customer_name.delete(0, tk.END); self.customer_name.insert(0, customer_data['name'])
        self.customer_reg_num.delete(0, tk.END); self.customer_reg_num.insert(0, customer_data['reg_num'])
        self.customer_address.delete(0, tk.END); self.customer_address.insert(0, customer_data['address'])
        self.customer_phone.delete(0, tk.END); self.customer_phone.insert(0, customer_data['phone'])
        self.customer_email.delete(0, tk.END); self.customer_email.insert(0, customer_data['email'])

        # messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_loaded"])

    def delete_customer_preset(self):
        preset_name = self.customer_preset_combobox.get()
        if not preset_name or preset_name not in self.customer_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        if messagebox.askyesno("Confirm Delete", self.i18n[self.language.get()]["confirm_delete_preset"]):
            del self.customer_presets[preset_name]
            self.save_customer_presets()
            self.update_customer_preset_combobox()
            messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_deleted"])

    def on_customer_preset_selected(self, event):
        selected_preset = self.customer_preset_combobox.get()
        self.customer_preset_name_entry.delete(0, tk.END)
        self.customer_preset_name_entry.insert(0, selected_preset)

        if selected_preset in self.customer_presets:
            customer_data = self.customer_presets[selected_preset]
            self.customer_name.delete(0, tk.END)
            self.customer_name.insert(0, customer_data.get('name', ''))
            self.customer_reg_num.delete(0, tk.END)
            self.customer_reg_num.insert(0, customer_data.get('reg_num', ''))
            self.customer_address.delete(0, tk.END)
            self.customer_address.insert(0, customer_data.get('address', ''))
            self.customer_phone.delete(0, tk.END)
            self.customer_phone.insert(0, customer_data.get('phone', ''))
            self.customer_email.delete(0, tk.END)
            self.customer_email.insert(0, customer_data.get('email', ''))

    

    def load_color_presets(self):
        preset_file = "color_presets.json"
        if os.path.exists(preset_file):
            try:
                with open(preset_file, "r", encoding="utf-8") as f:
                    self.color_presets = json.load(f)
            except json.JSONDecodeError:
                self.color_presets = {}
        self.update_color_preset_combobox()

    def save_color_presets(self):
        preset_file = "color_presets.json"
        with open(preset_file, "w", encoding="utf-8") as f:
            json.dump(self.color_presets, f, indent=4, ensure_ascii=False)

    def update_color_preset_combobox(self):
        self.color_preset_combobox["values"] = list(self.color_presets.keys())
        if self.color_presets:
            self.color_preset_combobox.set(list(self.color_presets.keys())[0])
        else:
            self.color_preset_combobox.set("")

    def remove_last_item_row(self):
        if self.items:
            last_item = self.items.pop()
            # 각 위젯을 개별적으로 제거합니다.
            last_item['name'].destroy()
            last_item['quantity'].destroy()
            last_item['unit_price'].destroy()
            last_item['amount'].destroy()
            self._schedule_preview_update()

    def add_item_row(self, index):
        # 각 Entry 위젯들을 self.items_frame에 직접 배치합니다.
        item_name = ttk.Entry(self.items_frame)
        item_name.grid(row=index + 1, column=0, padx=2, pady=2, sticky="we")
        item_name.bind("<KeyRelease>", lambda e: self._schedule_preview_update())

        item_quantity = ttk.Entry(self.items_frame)
        item_quantity.grid(row=index + 1, column=1, padx=2, pady=2, sticky="we")
        item_quantity.bind("<KeyRelease>", lambda e: self.calculate_amount(index))

        item_unit_price = ttk.Entry(self.items_frame)
        item_unit_price.grid(row=index + 1, column=2, padx=2, pady=2, sticky="we")
        item_unit_price.bind("<KeyRelease>", lambda e: self.calculate_amount(index))

        item_amount = ttk.Entry(self.items_frame, state="readonly")
        item_amount.grid(row=index + 1, column=3, padx=2, pady=2, sticky="we")

        self.items.append({
            # 'frame' 키는 더 이상 필요 없으므로 제거합니다.
            'name': item_name,
            'quantity': item_quantity,
            'unit_price': item_unit_price,
            'amount': item_amount
        })
        self.update_item_labels() # 품목 레이블 업데이트
        self._schedule_preview_update() # 미리보기 업데이트

    def calculate_amount(self, index):
        try:
            quantity = int(self.items[index]['quantity'].get() or 0)
            unit_price = int(self.items[index]['unit_price'].get() or 0)
            amount = quantity * unit_price
            self.items[index]['amount'].config(state=tk.NORMAL)
            self.items[index]['amount'].delete(0, tk.END)
            self.items[index]['amount'].insert(0, f"{amount:,}")
            self.items[index]['amount'].config(state="readonly")
            self._schedule_preview_update()
        except ValueError:
            self.items[index]['amount'].config(state=tk.NORMAL)
            self.items[index]['amount'].delete(0, tk.END)
            self.items[index]['amount'].insert(0, "")
            self.items[index]['amount'].config(state="readonly")
            # messagebox.showerror("Error", self.i18n[self.language.get()]["error_numeric"]) # 너무 자주 뜰 수 있으므로 주석 처리
            self._schedule_preview_update()

    

    def update_item_labels(self):
        lang = self.language.get()
        t = self.i18n[lang]
        # 품목 헤더 레이블 업데이트
        self.item_header_labels["item"].config(text=t["item"])
        self.item_header_labels["quantity"].config(text=t["quantity"])
        self.item_header_labels["unit_price"].config(text=t["unit_price"])
        self.item_header_labels["amount"].config(text=t["amount"])

    def add_color_preset(self):
        preset_name = self.color_preset_name_entry.get().strip()
        if not preset_name:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_name_empty"])
            return

        current_colors = {
            "primary_color": self.primary_color.get(),
            "secondary_color": self.secondary_color.get(),
            "text_color": self.text_color.get(),
            "light_text_color": self.light_text_color.get(),
            "border_color": self.border_color.get()
        }
        
        self.color_presets[preset_name] = current_colors
        self.save_color_presets()
        self.update_color_preset_combobox()
        messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_saved"])

    def apply_color_preset(self):
        preset_name = self.color_preset_combobox.get()
        if not preset_name or preset_name not in self.color_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        colors_data = self.color_presets[preset_name]
        self.primary_color.set(colors_data["primary_color"])
        self.secondary_color.set(colors_data["secondary_color"])
        self.text_color.set(colors_data["text_color"])
        self.light_text_color.set(colors_data["light_text_color"])
        self.border_color.set(colors_data["border_color"])

        # messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_loaded"])

    def delete_color_preset(self):
        preset_name = self.color_preset_combobox.get()
        if not preset_name or preset_name not in self.color_presets:
            messagebox.showerror("Error", self.i18n[self.language.get()]["preset_not_found"])
            return

        if messagebox.askyesno("Confirm Delete", self.i18n[self.language.get()]["confirm_delete_preset"]):
            del self.color_presets[preset_name]
            self.save_color_presets()
            self.update_color_preset_combobox()
            messagebox.showinfo("Success", self.i18n[self.language.get()]["preset_deleted"])

    def _create_pdf_data(self):
        lang = self.language.get()
        t = self.i18n[lang]
        data = {"lang": lang, "doc_type": t[self.doc_type.get()]}
        data["doc_number"] = self.doc_number.get()
        data["doc_date"] = self.doc_date.get()
        
        # 납부기한/유효기간 계산
        doc_date_obj = datetime.strptime(self.doc_date.get(), '%Y-%m-%d')
        due_date_obj = doc_date_obj + timedelta(days=30)
        data["due_date"] = due_date_obj.strftime('%Y-%m-%d')

        data["supplier"] = SUPPLIER_INFO[lang]
        
        data["customer"] = {
            "name": self.customer_name.get(),
            "reg_num": self.customer_reg_num.get(),
            "address": self.customer_address.get(),
            "phone": self.customer_phone.get(),
            "email": self.customer_email.get()
        }

        data["issuer"] = {
            "name": self.issuer_name.get(),
            "title": self.issuer_title.get(),
            "email": self.issuer_email.get(),
            "phone": self.issuer_phone.get()
        }
        
        items_data = []
        supply_amount = 0
        for item_widget in self.items:
            name = item_widget['name'].get()
            quantity_str = item_widget['quantity'].get()
            unit_price_str = item_widget['unit_price'].get()
            if name and quantity_str and unit_price_str:
                try:
                    quantity = int(quantity_str)
                    unit_price = int(unit_price_str)
                    amount = quantity * unit_price
                    supply_amount += amount
                    items_data.append({ 'name': name, 'quantity': quantity, 'unit_price': unit_price, 'amount': amount })
                except ValueError:
                    messagebox.showerror("Error", t["error_numeric"])
                    return None
        
        # if not data["customer"]["name"] or not items_data:
        #     messagebox.showerror("Error", t["error_fill_all"])
        #     return None

        data["items"] = items_data
        data["supply_amount"] = supply_amount
        if lang == 'ko':
            data["vat"] = int(supply_amount * 0.1)
            data["total_amount"] = data["supply_amount"] + data["vat"]
        else:
            data["vat"] = 0
            data["total_amount"] = supply_amount
        return data

    def _draw_pdf(self, buffer, doc_type_text, total_pages=1):
        data = self._create_pdf_data()
        if not data: return None

        lang, symbol = data["lang"], self.i18n[data["lang"]]["currency_symbol"]
        t = self.i18n[lang]
        
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4 # A4 사이즈로 변경

        # 일관된 내용 영역 정의
        content_x_start = 50
        content_width = width - (content_x_start * 2) # 양쪽 50pt 여백

        # --- 새로운 디자인 색상 및 글꼴 설정 ---
        primary_color = colors.HexColor(self.primary_color.get())
        secondary_color = colors.HexColor(self.secondary_color.get())
        text_color = colors.HexColor(self.text_color.get())
        light_text_color = colors.HexColor(self.light_text_color.get())
        border_color = colors.HexColor(self.border_color.get())

        # --- 1. 상단 헤더 --- 
        c.setFillColor(primary_color)
        c.rect(content_x_start, height - 60, content_width, 60, fill=1, stroke=0) # 높이 60으로 조정
        c.setFont(self.default_font_bold, 32)
        c.setFillColor(colors.HexColor(self.invoice_title_color.get())) # 텍스트 색상을 청구서 제목 색상으로 변경
        c.drawString(content_x_start + 20, height - 45, doc_type_text)

        # --- 2. 문서 번호, 작성일, 납부기한/유효기간, 발행자 정보 ---
        y_pos_doc_info_box_top = height - 80 # 조정된 헤더 높이에 맞춰 시작 위치 조정
        box_height = 70 # Adjusted height for 3 lines of text + padding
        c.setFillColor(secondary_color)
        c.rect(content_x_start, y_pos_doc_info_box_top - box_height, content_width, box_height, fill=1, stroke=0) # x, y (bottom-left), width, height
        
        y_pos_text_start = y_pos_doc_info_box_top - 15 # Start text slightly below box top

        c.setFont(self.default_font_bold, 9)
        c.setFillColor(text_color)
        c.drawString(content_x_start + 20, y_pos_text_start, t['doc_num'])
        c.drawString(content_x_start + 20, y_pos_text_start - 18, t['date'])
        
        if self.doc_type.get() == 'invoice':
            c.drawString(content_x_start + 20, y_pos_text_start - 36, t['due_date'])
        else:
            c.drawString(content_x_start + 20, y_pos_text_start - 36, t['validity_period'])

        c.drawString(content_x_start + content_width / 2 + 20, y_pos_text_start, t['issuer_info'])

        c.setFont(self.default_font, 9)
        c.drawString(content_x_start + 70, y_pos_text_start, data['doc_number']) # Adjusted X
        c.drawString(content_x_start + 70, y_pos_text_start - 18, data['doc_date']) # Adjusted X
        c.drawString(content_x_start + 70, y_pos_text_start - 36, data['due_date']) # This is the calculated due_date/validity_period

        c.drawString(content_x_start + content_width / 2 + 20, y_pos_text_start - 15, f"{data['issuer']['name']} ({data['issuer']['title']})")
        c.drawString(content_x_start + content_width / 2 + 20, y_pos_text_start - 30, data['issuer']['email'])
        c.drawString(content_x_start + content_width / 2 + 20, y_pos_text_start - 45, data['issuer']['phone'])

        # --- 3. 공급자 및 고객 정보 ---
        y_pos_supplier_customer_start = y_pos_doc_info_box_top - box_height - 20 # Gap of 20 points below doc info box
        c.setFont(self.default_font_bold, 12)
        c.setFillColor(text_color)
        c.drawString(content_x_start, y_pos_supplier_customer_start, t["from"])
        c.drawString(content_x_start + content_width / 2, y_pos_supplier_customer_start, t["to"])

        c.setStrokeColor(border_color)
        c.line(content_x_start, y_pos_supplier_customer_start - 10, content_x_start + content_width / 2 - 20, y_pos_supplier_customer_start - 10)
        c.line(content_x_start + content_width / 2, y_pos_supplier_customer_start - 10, content_x_start + content_width, y_pos_supplier_customer_start - 10)

        y_pos_supplier_customer_content_start = y_pos_supplier_customer_start - 30
        c.setFont(self.default_font, 10)
        # 공급자 정보
        c.drawString(content_x_start, y_pos_supplier_customer_content_start, data["supplier"]['name'])
        c.setFont(self.default_font, 9)
        c.setFillColor(colors.gray)
        c.drawString(content_x_start, y_pos_supplier_customer_content_start - 15, f"{t['reg_num']} {data['supplier']['reg_num']}")
        if lang == 'en':
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 30, data["supplier"]['address'][0])
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 45, data["supplier"]['address'][1])
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 60, f"{t['phone']} {data['supplier']['phone']}")
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 75, f"{t['email']} {data['supplier']['email']}")
        else:
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 30, data["supplier"]['address'])
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 45, f"{t['phone']} {data['supplier']['phone']}")
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 60, f"{t['email']} {data['supplier']['email']}")

        # 고객 정보
        y_pos_customer_content_start = y_pos_supplier_customer_start - 30 # Same starting Y as supplier content
        c.setFont(self.default_font, 10)
        c.setFillColor(text_color)
        c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start, f"{data['customer']['name']} {t['dear']}")
        c.setFont(self.default_font, 9)
        c.setFillColor(colors.gray)
        if lang == 'ko':
            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 15, f"{t['reg_num']} {data['customer']['reg_num']}")
            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 30, data['customer']['address'])
            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 45, f"{t['phone']} {data['customer']['phone']}")
            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 60, f"{t['email']} {data['customer']['email']}")
        else:
            # 영문 주소 줄바꿈 처리
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 15, f"{t['reg_num']} {data['supplier']['reg_num']}")
        if lang == 'en':
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 30, data["supplier"]['address'][0])
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 45, data["supplier"]['address'][1])
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 60, f"{t['phone']} {data['supplier']['phone']}")
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 75, f"{t['email']} {data['supplier']['email']}")
        else:
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 30, data["supplier"]['address'])
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 45, f"{t['phone']} {data['supplier']['phone']}")
            c.drawString(content_x_start, y_pos_supplier_customer_content_start - 60, f"{t['email']} {data['supplier']['email']}")

            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 15, data['customer']['address'])
            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 30, f"{t['phone']} {data['customer']['phone']}")
            c.drawString(content_x_start + content_width / 2, y_pos_customer_content_start - 45, f"{t['email']} {data['customer']['email']}")

        # --- 4. 품목 테이블 ---
        lowest_point_supplier_customer = y_pos_supplier_customer_content_start - 60 # Lowest text point
        y_pos_item_table = lowest_point_supplier_customer - 50 # 50 points gap below supplier/customer info (increased by 20)
        table_headers = [t["item"], t["quantity"], t["unit_price"], t["amount"]]
        col_widths = [285, 70, 70, 70] # Adjusted to sum to 512 (content_width)
        x_start = content_x_start

        c.setFillColor(primary_color)
        c.rect(x_start, y_pos_item_table - 5, content_width, 20, fill=1, stroke=0) # 높이 20으로 조정
        c.setFont(self.default_font_bold, 10)
        c.setFillColor(light_text_color)
        
        # 헤더 중앙 정렬
        current_x = x_start
        for i, header in enumerate(table_headers):
            c.drawCentredString(current_x + col_widths[i] / 2, y_pos_item_table, header)
            current_x += col_widths[i]

        y_pos_item_table_content_start = y_pos_item_table - 30
        c.setFont(self.default_font, 9)
        c.setFillColor(text_color)
        right_align_x = content_x_start + content_width - 10 # Define the right alignment x-coordinate
        for item in data["items"]:
            c.setStrokeColor(border_color)
            c.line(x_start, y_pos_item_table_content_start + 12, x_start + content_width, y_pos_item_table_content_start + 12)
            
            # 내용 정렬
            c.drawCentredString(x_start + col_widths[0] / 2, y_pos_item_table_content_start, item['name']) # Center align item name
            c.drawCentredString(x_start + col_widths[0] + col_widths[1] / 2, y_pos_item_table_content_start, f"{item['quantity']:,}") # Quantity remains centered
            if lang == 'ko':
                c.drawRightString(right_align_x - col_widths[3], y_pos_item_table_content_start, f"{item['unit_price']:,} {symbol}")
                c.drawRightString(right_align_x, y_pos_item_table_content_start, f"{item['amount']:,} {symbol}")
            else:
                c.drawRightString(right_align_x - col_widths[3], y_pos_item_table_content_start, f"{symbol}{item['unit_price']:,}")
                c.drawRightString(right_align_x, y_pos_item_table_content_start, f"{symbol}{item['amount']:,}")
            y_pos_item_table_content_start -= 25
        c.line(x_start, y_pos_item_table_content_start + 12, x_start + content_width, y_pos_item_table_content_start + 12)

        # --- 5. 합계 ---
        y_pos_summary_start = y_pos_item_table_content_start - 10 # Gap below item table
        total_x_start = content_x_start + content_width / 2 # Align with right half of content area
        if lang == 'ko':
            c.setFont(self.default_font, 10)
            c.drawString(total_x_start, y_pos_summary_start, t["supply_amount"])
            c.drawRightString(content_x_start + content_width - 10, y_pos_summary_start, f"{data['supply_amount']:,} {symbol}")
            y_pos_summary_start -= 20 # 간격 조정
            c.drawString(total_x_start, y_pos_summary_start, t["vat"])
            c.drawRightString(content_x_start + content_width - 10, y_pos_summary_start, f"{data['vat']:,} {symbol}")
            y_pos_summary_start -= 20 # 간격 조정
            c.setStrokeColor(border_color)
            c.line(total_x_start, y_pos_summary_start, content_x_start + content_width, y_pos_summary_start)
            y_pos_summary_start -= 5

        c.setFillColor(primary_color)
        c.rect(total_x_start - 20, y_pos_summary_start - 20, content_width / 2 + 20, 30, fill=1, stroke=0) # 높이 30으로 조정
        c.setFont(self.default_font_bold, 12)
        c.setFillColor(light_text_color)
        c.drawString(total_x_start, y_pos_summary_start - 10, t["total_amount"])
        if lang == 'ko':
            c.drawRightString(content_x_start + content_width - 10, y_pos_summary_start - 10, f"{data['total_amount']:,} {symbol}")
        else:
            c.drawRightString(content_x_start + content_width - 10, y_pos_summary_start - 10, f"{symbol}{data['total_amount']:,}")

        # --- Bank Fee Note (English only) ---
        if lang == 'en':
            fee_note_y_pos = y_pos_summary_start - 40 # Adjust Y position as needed
            c.setFont(self.default_font_bold, 9) # Smaller font for note, but bold
            c.setFillColor(text_color)
            c.drawString(content_x_start + content_width / 2, fee_note_y_pos, "All bank fees should be covered by the sender.")

        # --- 8. 하단 중앙 문구 ---
        c.setFont(self.default_font, 9)
        c.setFillColor(text_color)
        if lang == 'ko':
            footer_text = "궁금한 점 있으시면 편하게 연락해주세요."
            footer_y_pos = 110 # Above bank info for Korean
        else:
            footer_text = "Please feel free to contact us if you have any questions."
            footer_y_pos = 170 # Above bank info for English
        c.drawRightString(content_x_start + content_width - 10, footer_y_pos, footer_text) # Adjusted Y position and right-aligned

        # --- 6. 하단 정보 (계좌) ---
        # Move bank info to the bottom of the page
        if lang == 'ko':
            y_pos_bank_info_start = 150 # Adjusted for Korean
        else: # lang == 'en'
            y_pos_bank_info_start = 160 # Adjusted for English (more lines)

        c.setFont(self.default_font_bold, 12) # 글꼴 크기 키우고 볼드 처리
        c.setFillColor(text_color)
        c.drawString(content_x_start, y_pos_bank_info_start, t["bank_info"])
        c.setStrokeColor(border_color)
        c.line(content_x_start, y_pos_bank_info_start - 10, content_x_start + content_width / 2 - 20, y_pos_bank_info_start - 10)
        y_pos_bank_info_start -= 25
        c.setFont(self.default_font, 9)
        if lang == 'ko':
            c.drawString(content_x_start, y_pos_bank_info_start, "은행명: 국민은행")
            c.drawString(content_x_start, y_pos_bank_info_start - 15, "계좌번호: 823701-04-343660")
            c.drawString(content_x_start, y_pos_bank_info_start - 30, "예금주: 주식회사 애드캐리")
        else:
            c.drawString(content_x_start, y_pos_bank_info_start, "Name of Bank Holder: ADCARRY Corp.")
            c.drawString(content_x_start, y_pos_bank_info_start - 15, "Bank name: Citibank")
            c.drawString(content_x_start, y_pos_bank_info_start - 30, "Account Number: 73380000000193674")
            c.drawString(content_x_start, y_pos_bank_info_start - 45, "Routing Number: 031100209")
            c.drawString(content_x_start, y_pos_bank_info_start - 60, "Bank Address: 111 Wall Street, New york, New York 10043, United States")

        

        # --- 7. 페이지 번호 및 하단 줄 ---
        c.setStrokeColor(primary_color)
        c.setLineWidth(2) # 두꺼운 줄
        c.line(content_x_start, 30, content_x_start + content_width, 30) # 페이지 하단에 줄 추가

        c.setFont(self.default_font, 8)
        c.setFillColor(colors.gray)
        c.drawCentredString(content_x_start + content_width / 2, 15, f"Page {c.getPageNumber()}/{total_pages}")

        c.save()
        return data

    def generate_preview(self):
        if not PREVIEW_ENABLED: return
        self.pdf_buffer = io.BytesIO()
        doc_type_key = self.doc_type.get()
        lang = self.language.get()
        doc_type_text = self.i18n[lang][doc_type_key]
        
        pdf_data = self._draw_pdf(self.pdf_buffer, doc_type_text)
        if not pdf_data: self.pdf_buffer = None; return

        self.pdf_buffer.seek(0)
        pdf_document = fitz.open(stream=self.pdf_buffer.read(), filetype="pdf")
        total_pages = len(pdf_document) # 총 페이지 수 가져오기
        pdf_document.close()

        # 다시 PDF를 생성하여 total_pages를 전달
        self.pdf_buffer = io.BytesIO()
        pdf_data = self._draw_pdf(self.pdf_buffer, doc_type_text, total_pages=total_pages)
        if not pdf_data: self.pdf_buffer = None; return

        self.pdf_buffer.seek(0)
        pdf_document = fitz.open(stream=self.pdf_buffer.read(), filetype="pdf")
        self.preview_images = []
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            self.preview_images.append(img)
        pdf_document.close()
        self.current_page = 0
        self.update_preview_display()

    def update_preview_display(self):
        if not self.preview_images: return
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        img = self.preview_images[self.current_page]
        img_ratio = img.width / img.height
        canvas_ratio = canvas_width / canvas_height
        if img_ratio > canvas_ratio:
            new_width = canvas_width
            new_height = int(new_width / img_ratio)
        else:
            new_height = canvas_height
            new_width = int(new_height * img_ratio)
        resized_img = img.resize((new_width, new_height), Image.LANCZOS)
        self.photo_image = ImageTk.PhotoImage(resized_img)
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(canvas_width/2, canvas_height/2, anchor=tk.CENTER, image=self.photo_image)
        lang = self.language.get()
        t = self.i18n[lang]
        self.page_label.config(text=f"{t['page']} {self.current_page + 1} {t['of']} {len(self.preview_images)}")
        self.prev_button.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if self.current_page < len(self.preview_images) - 1 else tk.DISABLED)

    def show_previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_preview_display()

    def show_next_page(self):
        if self.current_page < len(self.preview_images) - 1:
            self.current_page += 1
            self.update_preview_display()

    def save_pdf_from_preview(self):
        lang = self.language.get()
        t = self.i18n[lang]
        if not self.pdf_buffer:
            messagebox.showerror("Error", t["error_preview"] + " (PDF buffer is empty)")
            print("DEBUG: PDF buffer is empty.")
            return
        save_dir = self.save_path.get()
        if not save_dir:
            messagebox.showerror("Error", t["error_path"])
            print("DEBUG: Save directory is empty.")
            return
        customer_name = self.customer_name.get()
        doc_type_key = self.doc_type.get()
        doc_type_text = self.i18n[lang][doc_type_key]

        # 파일 이름 정리 (특수 문자 제거)
        safe_customer_name = "".join(c for c in customer_name if c.isalnum() or c.isspace()).strip()
        if not safe_customer_name:
            safe_customer_name = "Untitled" # 고객 이름이 비어있을 경우 기본값

        safe_doc_type_text = "".join(c for c in doc_type_text if c.isalnum() or c.isspace()).strip()
        if not safe_doc_type_text:
            safe_doc_type_text = "Document" # 문서 종류가 비어있을 경우 기본값

        base_file_name = self.filename_var.get().strip()
        if not base_file_name:
            base_file_name = f"{safe_customer_name}_{safe_doc_type_text}"

        if self.add_date_prefix.get():
            date_prefix = datetime.now().strftime('%Y%m%d_')
            file_name = f"{date_prefix}{base_file_name}.pdf"
        else:
            file_name = f"{base_file_name}.pdf"
        file_path = os.path.join(save_dir, file_name)
        print(f"DEBUG: Attempting to save PDF to: {file_path}")
        try:
            with open(file_path, "wb") as f:
                f.write(self.pdf_buffer.getvalue())
            messagebox.showinfo("Success", f"{doc_type_text} {t['success_save']}\n{file_path}")
            print(f"DEBUG: PDF successfully saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"{t['error_save']}\n{e}")
            print(f"DEBUG: Error saving PDF: {e}")

if __name__ == '__main__':
    root = tk.Tk()
    if not PREVIEW_ENABLED:
        messagebox.showwarning("Dependency Missing", "PyMuPDF or Pillow is not installed.\nPreview functionality will be disabled.")
    app = InvoiceGenerator(root)
    root.after(100, app.generate_preview) # 100ms 지연 후 미리보기 생성
    root.mainloop()