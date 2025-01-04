import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
from datetime import datetime, timedelta
import sys
import os

class ProductManagement:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("产品出入库信息管理系统")
        self.root.geometry("1000x800")
        
        # 添加用户角色属性
        self.is_admin = False
        self.admin_password = "admin"  # 设置管理员密码
        
        # 设置窗口背景色
        self.root.configure(bg='#f0f2f5')
        
        # 初始化所有输入框变量为None
        self.id_entry = None
        self.name_entry = None
        self.price_entry = None
        self.colour_entry = None
        self.stock_entry = None
        self.date_entry = None
        self.remaining_stock_label = None
        self.product_tree = None
        
        # 添加排序状态属性
        self.order_desc = True  # True表示降序,False表示升序
        
        # 添加库存变动排序状态属性
        self.movement_order_desc = True  # True表示降序,False表示升序
        
        # 设置自定义样式
        self.setup_custom_styles()
        
        # 创建数据库连接
        self.create_database()
        
        # 在创建表之后，检查并添加 notes 列
        try:
            self.cursor.execute("""
                ALTER TABLE stock_movements ADD COLUMN notes TEXT
            """)
            self.conn.commit()
        except sqlite3.OperationalError:
            # 如果列已存在，忽略错误
            pass
        
        # 创建界面
        self.create_widgets()
        
        # 添加用户角色切换按钮
        self.create_role_switch_button()
        
    def create_database(self):
        try:
            # 获取用户文档目录
            user_docs = os.path.expanduser('~/Documents')
            # 创建应用程序数据目录
            app_data_dir = os.path.join(user_docs, 'ProductManagement')
            
            # 如果目录不存在则创建
            if not os.path.exists(app_data_dir):
                os.makedirs(app_data_dir)
            # 设置数据库路径
            db_path = os.path.join(app_data_dir, 'store.db')
            self.conn = sqlite3.connect(db_path)
            self.cursor = self.conn.cursor()
            
            # 创建新的商品表
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS products (
                    product_id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    price REAL NOT NULL,
                    colour TEXT NOT NULL,
                    stock INTEGER NOT NULL,
                    order_date TEXT
                )
            ''')
            
            # 添加创建库存变动表的SQL语句
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS stock_movements (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    product_id TEXT NOT NULL,
                    movement_type TEXT NOT NULL,
                    quantity INTEGER NOT NULL,
                    datetime TEXT NOT NULL,
                    remarks TEXT,
                    notes TEXT,
                    FOREIGN KEY (product_id) REFERENCES products(product_id)
                )
            ''')
            
            self.conn.commit()
        except Exception as e:
            messagebox.showerror("错误", f"创建数据库失败：{str(e)}")

    def create_widgets(self):
        # 设置整体样式
        style = ttk.Style()
        style.configure("TNotebook", padding=5)
        style.configure("TFrame", padding=5)
        style.configure("TLabelframe", padding=10)
        style.configure("TButton", padding=5)
        
        # 创建notebook并设置样式
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 修改标签样式，使其更加醒目
        style = ttk.Style()
        style.configure("Tab.TFrame", background="#ffffff")
        style.configure("TNotebook.Tab", 
            padding=[20, 8],           # 增加内边距使标签更大
            font=('Microsoft YaHei', 11, 'bold'),  # 使用更大更粗的字体
            background="#e3f2fd",      # 设置背景色为浅蓝色
            foreground="#1976d2"       # 设置文字颜色为深蓝色
        )
        style.map("TNotebook.Tab",
            background=[("selected", "#bbdefb")],  # 选中时的背景色
            foreground=[("selected", "#0d47a1")],  # 选中时的文字颜色
            expand=[("selected", [1, 1, 1, 0])]    # 选中时的扩展效果
        )
        
        # 商品管理页面
        self.product_frame = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.product_frame, text='  产品管理  ')
        
        # 库存变动页面
        self.stock_frame = ttk.Frame(self.notebook, style="Tab.TFrame")
        self.notebook.add(self.stock_frame, text='  库存变动  ')
        
        self.create_product_page()
        self.create_stock_page()
    
    def create_product_page(self):
        # 创建水平布局的主框架
        main_frame = ttk.Frame(self.product_frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 左侧产品信息框架 - 设置固定宽度比例
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # 商品信息输入框架
        input_frame = ttk.LabelFrame(left_frame, text="产品信息", padding=10)
        input_frame.pack(fill="x", pady=5)
        
        # 创建输入框
        ttk.Label(input_frame, text="产品型号:", style='Label.TLabel').grid(row=0, column=0, padx=5, pady=5)
        self.id_entry = ttk.Entry(input_frame, style='Entry.TEntry')
        self.id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(input_frame, text="产品甲方:", style='Label.TLabel').grid(row=0, column=2, padx=5, pady=5)
        self.name_entry = ttk.Entry(input_frame, style='Entry.TEntry')
        self.name_entry.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(input_frame, text="产品价格:", style='Label.TLabel').grid(row=1, column=0, padx=5, pady=5)
        self.price_entry = ttk.Entry(input_frame, style='Entry.TEntry')
        self.price_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(input_frame, text="产品颜色:", style='Label.TLabel').grid(row=1, column=2, padx=5, pady=5)
        self.colour_entry = ttk.Entry(input_frame, style='Entry.TEntry')
        self.colour_entry.grid(row=1, column=3, padx=5, pady=5)
        
        # 订单数量和剩余库存放在同一行
        ttk.Label(input_frame, text="订单数量:", style='Label.TLabel').grid(row=2, column=0, padx=5, pady=5)
        self.stock_entry = ttk.Entry(input_frame, style='Entry.TEntry')
        self.stock_entry.grid(row=2, column=1, padx=5, pady=5)
        
        # 剩余库存放在第二列
        ttk.Label(input_frame, text="剩余库存:", style='Label.TLabel').grid(row=2, column=2, padx=5, pady=5)
        self.remaining_stock_label = ttk.Label(input_frame, 
                                         text="0",
                                         font=('Microsoft YaHei', 10, 'bold'),  # 使用微软雅黑字体
                                         foreground='#007bff',           # 使用主题蓝色
                                         anchor='w')                  
        self.remaining_stock_label.grid(row=2, column=3, padx=(5,20), pady=5, sticky='w')  # 增加右侧padding
        
        # 订单日期和出库金额放在同一行
        ttk.Label(input_frame, text="订单日期:", style='Label.TLabel').grid(row=3, column=0, padx=5, pady=5)
        self.date_entry = ttk.Entry(input_frame, style='Entry.TEntry')
        self.date_entry.grid(row=3, column=1, padx=5, pady=5)
        
        ttk.Label(input_frame, text="出库金额:", style='Label.TLabel').grid(row=3, column=2, padx=5, pady=5)
        self.outbound_amount_label = ttk.Label(input_frame,
                                         text="¥0",  # 修改初始值，去掉小数位
                                         font=('Microsoft YaHei', 10, 'bold'),
                                         foreground='#dc3545',  # 使用红色显示金额
                                         anchor='w')
        self.outbound_amount_label.grid(row=3, column=3, padx=5, pady=5, sticky='w')
        
        # 在日期输入框旁添加按钮，调整按钮顺序
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=4, column=0, columnspan=4, padx=5, pady=5)
        
        # 查询产品按钮
        ttk.Button(button_frame, text="查询产品", 
                  command=self.search_product,
                  style='secondary.TButton').pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="增加产品", 
                  command=self.add_product,
                  style='primary.TButton').pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="重置", 
                  command=self.reset_product_form,
                  style='secondary.TButton').pack(side="left", padx=5)
        
        # 上传产品按钮
        ttk.Button(button_frame, text="上传产品",
                  command=self.upload_products,
                  style='info.TButton').pack(side="left", padx=5)
        
        # 下载模板按钮移到上传产品右边
        ttk.Button(button_frame, text="下载模板", 
                  command=self.download_template,
                  style='secondary.TButton').pack(side="left", padx=5)
        
        # 按钮框架
        button_frame = ttk.Frame(self.product_frame)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        # 修改按钮列表,移除增加产品按钮
        buttons = [
            ("修改产品", self.update_product, 'info', False),
            ("删除产品", self.delete_product, 'danger', False),
            ("库存入口", self.open_stock_page, 'secondary', True),  # 修改这里
            ("显示所有", self.show_all_products, 'secondary', True),
            ("快速出货", self.quick_stock_out, 'warning', True),
            ("快速入库", self.quick_stock_in, 'success', True),
            ("时间↓", self.generate_order, 'success', True),
            ("重置", self.reset_product_form, 'secondary', True)
        ]
        
        # 存储需要管理员权限的按钮引用
        self.admin_buttons = []
        
        for text, command, style, enabled in buttons:
            btn = ttk.Button(button_frame, text=text, command=command, style=f'{style}.TButton')
            if not enabled:
                btn.configure(state='disabled')
                if text in ["删除产品", "修改产品"]:  # 添加修改产品到管理员权限
                    self.admin_buttons.append(btn)
            btn.pack(side="left", padx=5)
        
        # 右侧高级查询框架 - 设置固定宽度
        right_frame = ttk.Frame(main_frame, width=300)  # 设置固定宽度
        right_frame.pack(side="left", fill="both", padx=(5, 0))
        right_frame.pack_propagate(False)  # 防止框架被内容压缩
        
        search_frame = ttk.LabelFrame(right_frame, text="高级查询", padding=10)
        search_frame.pack(fill="both", expand=True, pady=5)
        
        # 调整查询条件的网格布局
        ttk.Label(search_frame, text="产品型号:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.search_id_entry = ttk.Entry(search_frame, width=15)  # 设置输入框宽度
        self.search_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(search_frame, text="产品甲方:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.search_name_entry = ttk.Entry(search_frame, width=15)
        self.search_name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(search_frame, text="开始日期:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.search_start_date = ttk.Entry(search_frame, width=15)
        self.search_start_date.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="选择", 
                  command=lambda: self.search_start_date.delete(0, tk.END) or 
                                self.search_start_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
                  ).grid(row=2, column=2, padx=1, pady=5)
        
        ttk.Label(search_frame, text="结束日期:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        self.search_end_date = ttk.Entry(search_frame, width=15)
        self.search_end_date.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="选择", 
                  command=lambda: self.search_end_date.delete(0, tk.END) or 
                                self.search_end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
                  ).grid(row=3, column=2, padx=1, pady=5)
        
        # 查询按钮框架
        button_frame = ttk.Frame(search_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="高级查询", 
                  command=self.advanced_search,
                  style='info.TButton').pack(side="left", padx=5)
                  
        ttk.Button(button_frame, text="导出查询结果", 
                  command=self.export_search_results,
                  style='success.TButton').pack(side="left", padx=5)
        
        # 商品列表框架 (移到主框架下方)
        list_frame = ttk.Frame(self.product_frame)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 商品列表
        self.product_tree = ttk.Treeview(list_frame, 
                                    columns=("id", "name", "price", "colour", "stock", "date", "remaining", "outbound_amount"), 
                                    show="headings",
                                    style='Custom.Treeview')
        
        # 设置表头并绑定排序事件
        columns = [
            ("id", "产品型号", 100),
            ("name", "产品甲方", 150),
            ("price", "产品价格", 100),
            ("colour", "产品颜色", 100),
            ("stock", "订单数量", 100),
            ("date", "订单日期", 100),
            ("remaining", "剩余库存", 100),
            ("outbound_amount", "出库金额", 100)
        ]
        
        # 添加排序状态字典
        self.sort_states = {}
        for col in ["product_id", "name", "price", "colour", "stock", "order_date", "remaining_stock", "outbound_amount"]:
            self.sort_states[col] = True  # True表示降序
        
        for col, text, width in columns:
            self.product_tree.heading(col, text=text, anchor="center",
                                    command=lambda c=col: self.sort_treeview(c))
            self.product_tree.column(col, width=width, anchor="center")
        
        # 绑定事件
        self.product_tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.product_tree.bind('<FocusOut>', self.on_tree_focus_out)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.product_tree.yview)
        self.product_tree.configure(yscrollcommand=scrollbar.set)
        
        # 打包到界面
        self.product_tree.pack(fill="both", expand=True, padx=10, pady=5, side="left")
        scrollbar.pack(fill="y", pady=5, side="right")
    
    def create_stock_page(self):
        # 主框架使用更紧凑的布局
        main_frame = ttk.Frame(self.stock_frame)
        main_frame.pack(fill="both", expand=True)
        
        # 左侧输入和按钮区域 - 增加宽度
        left_frame = ttk.Frame(main_frame, width=550)  # 增加宽度
        left_frame.pack(side="left", fill="both", padx=20)  # 增加左侧框架的外边距
        left_frame.pack_propagate(False)  # 防止自动收缩
        
        # 库存变动登记框架 - 调整内边距
        input_frame = ttk.LabelFrame(left_frame, text="库存变动登记", padding=10)  # 增加内边距
        input_frame.pack(fill="both", expand=True, pady=5)  # 使用expand=True填充空间
        
        # 输入区域使用网格布局 - 增加间距
        ttk.Label(input_frame, text="产品型号:").grid(row=0, column=0, padx=5, pady=8, sticky='e')  # 增加间距
        self.movement_id_entry = ttk.Entry(input_frame, width=15)  # 增加输入框宽度
        self.movement_id_entry.grid(row=0, column=1, padx=5, pady=8)
        
        ttk.Label(input_frame, text="变动类型:").grid(row=0, column=2, padx=5, pady=8, sticky='e')
        self.movement_type = ttk.Combobox(input_frame, values=["入库", "出库"], width=12)  # 增加下拉框宽度
        self.movement_type.grid(row=0, column=3, padx=5, pady=8)
        
        ttk.Label(input_frame, text="变动数量:").grid(row=1, column=0, padx=5, pady=8, sticky='e')
        self.quantity_entry = ttk.Entry(input_frame, width=15)  # 增加输入框宽度
        self.quantity_entry.grid(row=1, column=1, padx=5, pady=8)
        
        # 库存日期输入框和今天按钮
        date_frame = ttk.Frame(input_frame)
        date_frame.grid(row=1, column=2, columnspan=2, padx=5, pady=8, sticky='w')
        
        ttk.Label(date_frame, text="库存日期:").pack(side='left', padx=(0,5))
        self.remarks_entry = ttk.Entry(date_frame, width=12)
        self.remarks_entry.pack(side='left', padx=(0,5))
        
        # 添加"今天"按钮
        ttk.Button(date_frame, text="今天", 
                  command=lambda: self.set_today_date(),
                  width=6).pack(side='left')
        
        # 默认填入今天日期
        self.remarks_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # 添加日期格式验证
        validate_cmd = (self.root.register(self.validate_date), '%P')
        self.remarks_entry.config(validate='key', validatecommand=validate_cmd)
        
        # 备注说明
        ttk.Label(input_frame, text="备注说明:").grid(row=2, column=0, padx=5, pady=8, sticky='ne')
        self.notes_text = tk.Text(input_frame, width=35, height=2)  # 调整文本框宽度
        self.notes_text.grid(row=2, column=1, columnspan=2, padx=5, pady=8, sticky='w')
        
        # 按钮区域
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=3, padx=5, pady=8, sticky='sw')  # 将按钮框架移到第2行第3列
        
        # 记录变动按钮
        ttk.Button(button_frame, text="记录变动",
                  command=self.record_stock_movement,
                  style='primary.TButton').pack(side="left")
        
        # 右侧查询区域 - 保持宽度
        right_frame = ttk.Frame(main_frame, width=350)
        right_frame.pack(side="left", fill="both", padx=(30, 30))  # 增加左右边距，特别是左边距
        right_frame.pack_propagate(False)
        
        # 记录查询框架
        search_frame = ttk.LabelFrame(right_frame, text="记录查询", padding=10)
        search_frame.pack(fill="both", expand=True, pady=5)
        
        # 查询条件布局
        ttk.Label(search_frame, text="产品型号:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.search_product_entry = ttk.Entry(search_frame, width=15)
        self.search_product_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(search_frame, text="开始日期:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.start_date_entry = ttk.Entry(search_frame, width=15)
        self.start_date_entry.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="选择",
                  command=lambda: self.start_date_entry.delete(0, tk.END) or 
                                self.start_date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
                  ).grid(row=1, column=2, padx=1, pady=5)
        
        ttk.Label(search_frame, text="结束日期:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.end_date_entry = ttk.Entry(search_frame, width=15)
        self.end_date_entry.grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(search_frame, text="选择",
                  command=lambda: self.end_date_entry.delete(0, tk.END) or 
                                self.end_date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
                  ).grid(row=2, column=2, padx=1, pady=5)
        
        # 查询按钮框架
        button_frame = ttk.Frame(search_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="查询记录",
                  command=self.search_movements,
                  style='info.TButton').pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="导出查询结果",
                  command=self.export_current_view,
                  style='success.TButton').pack(side="left", padx=5)
        
        # 操作按钮框架
        button_frame = ttk.Frame(self.stock_frame)
        button_frame.pack(fill="x", padx=5, pady=2)
        
        buttons = [
            ("修改记录", self.update_movement, 'info', False),
            ("删除记录", self.delete_movement, 'danger', False),
            ("产品入口", self.open_product_page, 'info', True),
            ("显示所有", self.view_all_movements, 'info', True),
            ("重置", self.reset_movement_form, 'warning', True),  # 将"导出当前"改为"重置"
            ("导出全部", self.export_movement_report, 'success', False),
            ("高级查询", self.advanced_movement_search, 'info', True),  # 将"清除历史"改为"高级查询"
            ("时间↓", self.sort_movements_by_time, 'success', True),
        ]
        
        # 存储需要管理员权限的按钮引用
        if not hasattr(self, 'admin_buttons'):
            self.admin_buttons = []
        
        for text, command, style, enabled in buttons:
            # 跳过变动时间按钮
            if text == "变动时间":
                continue
            btn = ttk.Button(button_frame,
                           text=text,
                           command=command,
                           style=f'{style}.TButton')
            if not enabled:
                btn.configure(state='disabled')
                if text in ["删除记录", "清除历史", "修改记录", "导出全部"]:
                    self.admin_buttons.append(btn)
            btn.pack(side="left", padx=2, pady=2)
        
        # 修改变动记录列表的创建和配置
        self.movement_tree = ttk.Treeview(self.stock_frame,
                                        columns=("id", "type", "quantity", "stock_date", "datetime", "notes"),
                                        show="headings",
                                        height=15)
        
        # 设置列宽和标题，添加排序功能
        columns = [
            ("id", "产品型号", 120),
            ("type", "变动类型", 80),
            ("quantity", "变动数量", 80),
            ("stock_date", "库存日期", 100),
            ("datetime", "变动时间", 120),
            ("notes", "备注说明", 200)
        ]
        
        # 添加排序状态字典
        self.movement_sort_states = {}
        for col, _, _ in columns:
            self.movement_sort_states[col] = True  # True表示降序
        
        # 设置列配置
        for col, text, width in columns:
            self.movement_tree.heading(col, text=text, anchor="center",
                                     command=lambda c=col: self.sort_movement_tree(c))
            self.movement_tree.column(col, width=width, anchor="center")
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.stock_frame, orient="vertical", command=self.movement_tree.yview)
        self.movement_tree.configure(yscrollcommand=scrollbar.set)
        
        # 移动状态栏到树形视图之前
        # 创建库存变动状态栏
        self.stock_status_label = ttk.Label(
            self.stock_frame,
            text="准备就绪",
            relief="sunken",
            padding=(5, 2)
        )
        self.stock_status_label.pack(side="bottom", fill="x", padx=5, pady=2)
        
        # 打包列表和滚动条
        self.movement_tree.pack(side="left", fill="both", expand=True, padx=5, pady=2)
        scrollbar.pack(side="right", fill="y", pady=2)
        
        # 绑定选择事件
        self.movement_tree.bind('<<TreeviewSelect>>', self.on_movement_tree_select)
    
    def record_stock_movement(self):
        try:
            product_id = self.movement_id_entry.get()
            movement_type = self.movement_type.get()
            quantity = int(self.quantity_entry.get())
            remarks = self.remarks_entry.get()
            notes = self.notes_text.get("1.0", tk.END).strip()  # 获取备注说明内容
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 检查商品是否存在
            self.cursor.execute("SELECT stock FROM products WHERE product_id=?", (product_id,))
            result = self.cursor.fetchone()
            if not result:
                messagebox.showerror("错误", "商品不存在！")
                return
            
            current_stock = result[0]
            
            # 计算库存变化
            if movement_type == "出库":
                if current_stock < quantity:
                    messagebox.showerror("错误", "库存不足！")
                    return
                quantity_change = -quantity
            else:  # 入库
                quantity_change = quantity
            
            # 更新库存
            self.cursor.execute("""
                UPDATE products 
                SET stock = stock + ? 
                WHERE product_id = ?""", (quantity_change, product_id))
            
            # 记录变动 - 调整插入顺序以匹配列顺序
            self.cursor.execute("""
                INSERT INTO stock_movements 
                (product_id, movement_type, quantity, datetime, remarks, notes)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (product_id, movement_type, quantity, 
                 current_time,
                 remarks, notes))
            
            self.conn.commit()
            messagebox.showinfo("成功", "库存变动已记录！")
            self.clear_movement_entries()
            self.show_all_products()
            self.view_stock_movements()
            
        except Exception as e:
            messagebox.showerror("错误", f"记录失败：{str(e)}")
    
    def view_stock_movements(self):
        """保持向后兼容,调用查看所有记录的方法"""
        self.view_all_movements()
    
    def export_movement_report(self):
        """导出全部库存变动记录"""
        try:
            self.cursor.execute("""
                SELECT product_id, movement_type, quantity, datetime, remarks
                FROM stock_movements
                ORDER BY datetime DESC
            """)
            
            data = self.cursor.fetchall()
            if not data:
                messagebox.showinfo("提示", "没有可导出的数据")
                return
            
            import pandas as pd
            df = pd.DataFrame(data, columns=['商品编号', '变动类型', '数量', '库存日期', '变动时间'])  # 修改这里的"出库日期"为"库存日期"
            
            # 生成默认文件名
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"出入库明细表_{current_time}.xlsx"
            
            # 打开文件选择对话框
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if filename:  # 如果用户选择了保存路径
                df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"报表已导出至：{filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
    
    def quick_stock_out(self):
        """优化的快速出货功能"""
        # 获取选中的商品
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要出货的商品！")
            return
        
        item = self.product_tree.item(selected[0])
        values = item['values']
        product_id = values[0]
        remaining_stock = values[6]
        
        # 保存当前查询条件
        current_search = {
            'product_id': self.search_id_entry.get().strip(),
            'name': self.search_name_entry.get().strip(),
            'start_date': self.search_start_date.get().strip(),
            'end_date': self.search_end_date.get().strip()
        }
        
        # 创建快速出货对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("快速出货")
        dialog.geometry("300x200")
        dialog.transient(self.root)  # 设置为主窗口的子窗口
        dialog.grab_set()  # 模态对话框
        dialog.resizable(False, False)  # 禁止调整大小
        
        # 计算居中位置
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        dialog_width = 300
        dialog_height = 200
        
        x = root_x + (root_width - dialog_width) // 2
        y = root_y + (root_height - dialog_height) // 2
        
        # 设置对话框位置
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 显示商品信息
        info_frame = ttk.Frame(dialog)
        info_frame.pack(pady=10)
        
        ttk.Label(info_frame, text=f"商品型号: {product_id}").pack()
        ttk.Label(info_frame, text=f"剩余库存: {remaining_stock}").pack()
        
        # 创建输入框架
        input_frame = ttk.Frame(dialog)
        input_frame.pack(pady=10)
        
        # 出货数量行
        quantity_frame = ttk.Frame(input_frame)
        quantity_frame.pack(fill='x', pady=5)
        ttk.Label(quantity_frame, text="出货数量:").pack(side='left', padx=5)
        quantity_entry = ttk.Entry(quantity_frame, width=15)
        quantity_entry.pack(side='left', padx=5)
        
        # 出货日期行
        date_frame = ttk.Frame(input_frame)
        date_frame.pack(fill='x', pady=5)
        ttk.Label(date_frame, text="出货日期:").pack(side='left', padx=5)
        date_entry = ttk.Entry(date_frame, width=15)
        date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        date_entry.pack(side='left', padx=5)
        
        def confirm_stock_out():
            try:
                quantity = int(quantity_entry.get())
                if quantity <= 0:
                    messagebox.showerror("错误", "出货数量必须大于0")
                    return
                    
                if quantity > int(remaining_stock):
                    messagebox.showerror("错误", "出货数量不能大于剩余库存")
                    return
                    
                # 记录出货
                self.cursor.execute("""
                    INSERT INTO stock_movements 
                    (product_id, movement_type, quantity, remarks, datetime, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (product_id, "出库", quantity, 
                      date_entry.get(),  # 出库日期
                      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # 变动时间
                      self.notes_text.get("1.0", tk.END).strip()))  # 备注说明
                      
                self.conn.commit()
                dialog.destroy()
                
                # 如果有查询条件，重新执行查询
                if any(current_search.values()):
                    # 恢复查询条件
                    self.search_id_entry.delete(0, tk.END)
                    self.search_id_entry.insert(0, current_search['product_id'])
                    
                    self.search_name_entry.delete(0, tk.END)
                    self.search_name_entry.insert(0, current_search['name'])
                    
                    self.search_start_date.delete(0, tk.END)
                    self.search_start_date.insert(0, current_search['start_date'])
                    
                    self.search_end_date.delete(0, tk.END)
                    self.search_end_date.insert(0, current_search['end_date'])
                    
                    # 重新执行查询
                    self.advanced_search()
                else:
                    # 如果没有查询条件，显示所有产品
                    self.show_all_products()
                
                # 重新选中之前的产品
                for item in self.product_tree.get_children():
                    if self.product_tree.item(item)['values'][0] == product_id:
                        self.product_tree.selection_set(item)
                        self.product_tree.see(item)
                        break
                
                self.view_all_movements()
                messagebox.showinfo("成功", f"已完成{quantity}件商品的出货操作")
                
            except ValueError:
                messagebox.showerror("错误", "请输入有效的出货数量")
        
        # 确认按钮
        ttk.Button(dialog, text="确认出货", command=confirm_stock_out).pack(pady=10)
        
        # 绑定回车键
        quantity_entry.bind('<Return>', lambda e: confirm_stock_out())
        
        # 设置初始焦点
        quantity_entry.focus()

    def clear_movement_entries(self):
        self.movement_id_entry.delete(0, tk.END)
        self.movement_type.set("")
        self.quantity_entry.delete(0, tk.END)
        self.remarks_entry.delete(0, tk.END)
        self.notes_text.delete("1.0", tk.END)  # 清空备注说明
    
    def clear_movement_tree(self):
        for item in self.movement_tree.get_children():
            self.movement_tree.delete(item)

    def add_product(self):
        """添加新产品"""
        try:
            # 获取输入的产品ID
            product_id = self.id_entry.get()
            
            # 检查产品ID是否已存在
            self.cursor.execute("SELECT * FROM products WHERE product_id=?", (product_id,))
            if self.cursor.fetchone():
                messagebox.showwarning("警告", "已存在相同的产品型号！")
                # 清空所有输入框
                self.clear_entries()
                return
            
            # 如果不存在相同ID，获取其他输入值
            name = self.name_entry.get()
            price = float(self.price_entry.get())
            colour = self.colour_entry.get()
            stock = int(self.stock_entry.get())
            
            # 如果日期为空，自动生成当前日期
            order_date = self.date_entry.get()
            if not order_date:
                order_date = datetime.now().strftime("%Y-%m-%d")
                self.date_entry.insert(0, order_date)
            
            # 插入新产品
            self.cursor.execute("""
                INSERT INTO products (product_id, name, price, colour, stock, order_date)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (product_id, name, price, colour, stock, order_date))
            
            self.conn.commit()
            messagebox.showinfo("成功", "产品添加成功！")
            self.clear_entries()
            self.show_all_products()
            self.update_status_bar()  # 更新状态栏
            
        except Exception as e:
            messagebox.showerror("错误", f"添加失败：{str(e)}")

    def update_product(self):
        try:
            product_id = self.id_entry.get()
            name = self.name_entry.get()
            price = float(self.price_entry.get())
            colour = self.colour_entry.get()  # 获取颜色值
            stock = int(self.stock_entry.get())
            order_date = self.date_entry.get()  # 获取日期值
            
            self.cursor.execute("""
                UPDATE products 
                SET name=?, price=?, colour=?, stock=?, order_date=?
                WHERE product_id=?
            """, (name, price, colour, stock, order_date, product_id))
            
            if self.cursor.rowcount == 0:
                messagebox.showerror("错误", "商品不存在！")
                return
            
            self.conn.commit()
            messagebox.showinfo("成功", "商品更新成功！")
            self.clear_entries()
            self.show_all_products()
            self.update_status_bar()  # 更新状态栏
            
        except Exception as e:
            messagebox.showerror("错误", f"更新失败：{str(e)}")
    
    def delete_product(self):
        product_id = self.id_entry.get()
        if messagebox.askyesno("确认", "确定要删除该商品吗？"):
            try:
                self.cursor.execute("DELETE FROM products WHERE product_id=?", (product_id,))
                if self.cursor.rowcount == 0:
                    messagebox.showerror("错误", "商品不存在！")
                    return
                
                self.conn.commit()
                messagebox.showinfo("成功", "商品删除成功！")
                self.clear_entries()
                self.show_all_products()
                self.update_status_bar()  # 更新状态栏
                
            except Exception as e:
                messagebox.showerror("错误", f"删除失败：{str(e)}")
    
    def search_product(self):
        product_id = self.id_entry.get()
        try:
            # 修改SQL查询，包含剩余库存计算
            self.cursor.execute("""
                SELECT p.product_id, p.name, p.price, p.colour, p.stock,
                       p.order_date,
                       (p.stock - COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           WHEN m.movement_type = '入库' THEN -m.quantity 
                           ELSE 0 
                       END), 0)) as remaining_stock,
                       (COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           ELSE 0 
                       END), 0) * p.price) as outbound_amount
                    FROM products p
                    LEFT JOIN stock_movements m ON p.product_id = m.product_id
                    WHERE p.product_id=?
                    GROUP BY p.product_id, p.name, p.price, p.colour, p.stock, p.order_date
            """, (product_id,))
            
            self.clear_tree()
            product = self.cursor.fetchone()
            if product:
                # 插入数据到树形视图
                self.product_tree.insert("", "end", values=product)
                
                # 自动填充输入框
                self.id_entry.delete(0, tk.END)
                self.id_entry.insert(0, product[0])  # product_id
                
                self.name_entry.delete(0, tk.END)
                self.name_entry.insert(0, product[1])  # name
                
                self.price_entry.delete(0, tk.END)
                self.price_entry.insert(0, product[2])  # price
                
                self.colour_entry.delete(0, tk.END)
                self.colour_entry.insert(0, product[3])  # colour
                
                self.stock_entry.delete(0, tk.END)
                self.stock_entry.insert(0, product[4])  # stock
                
                self.date_entry.delete(0, tk.END)
                self.date_entry.insert(0, product[5])  # order_date
                
                # 更新剩余库存显示
                self.remaining_stock_label.config(text=str(product[6]))  # remaining_stock
                
                # 更新出库金额显示
                outbound_amount = product[7] if len(product) > 7 else 0
                self.outbound_amount_label.config(text=f"¥{int(outbound_amount)}")  # 去掉小数位
            else:
                messagebox.showinfo("提示", "未找到该商品！")
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")
            print(f"Debug - Error: {str(e)}")  # 添加调试信息
    
    def show_all_products(self):
        self.clear_tree()
        try:
            # 修改查询以计算剩余库存，限制显示最近30天的记录
            self.cursor.execute("""
                SELECT p.product_id, p.name, p.price, p.colour, p.stock,
                       p.order_date,
                       (p.stock - COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           WHEN m.movement_type = '入库' THEN -m.quantity 
                           ELSE 0 
                       END), 0)) as remaining_stock,
                       (COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           ELSE 0 
                       END), 0) * p.price) as outbound_amount
                FROM products p
                LEFT JOIN stock_movements m ON p.product_id = m.product_id
                WHERE p.order_date >= date('now', '-30 day')
                GROUP BY p.product_id, p.name, p.price, p.colour, p.stock, p.order_date
                ORDER BY p.order_date DESC
            """)
            
            recent_products = self.cursor.fetchall()
            for product in recent_products:
                # 将最后一个值（出库金额）转为整数
                values = list(product)
                if values[7] is not None:  # 确保出库金额不是 None
                    values[7] = int(values[7])
                self.product_tree.insert("", "end", values=values)
            
            # 更新状态栏
            self.update_status_bar(len(recent_products))
            
            # 重置所有列的排序状态和表头显示
            for col in self.product_tree["columns"]:
                self.sort_states[col] = True
                text = self.product_tree.heading(col)["text"]
                if "↓" in text or "↑" in text:
                    text = text.rstrip("↓↑")
                self.product_tree.heading(col, text=text)
            
        except Exception as e:
            messagebox.showerror("错误", f"显示失败：{str(e)}")

    def clear_entries(self):
        # 检查输入框是否存在
        if hasattr(self, 'id_entry') and self.id_entry:
            self.id_entry.delete(0, tk.END)
        if hasattr(self, 'name_entry') and self.name_entry:
            self.name_entry.delete(0, tk.END)
        if hasattr(self, 'price_entry') and self.price_entry:
            self.price_entry.delete(0, tk.END)
        if hasattr(self, 'colour_entry') and self.colour_entry:
            self.colour_entry.delete(0, tk.END)
        if hasattr(self, 'stock_entry') and self.stock_entry:
            self.stock_entry.delete(0, tk.END)
        if hasattr(self, 'date_entry') and self.date_entry:
            self.date_entry.delete(0, tk.END)
        if hasattr(self, 'remaining_stock_label') and self.remaining_stock_label:
            self.remaining_stock_label.config(text="0")
        if hasattr(self, 'outbound_amount_label') and self.outbound_amount_label:
            self.outbound_amount_label.config(text="¥0")  # 修改初始值，去掉小数位
    
    def clear_tree(self):
        for item in self.product_tree.get_children():
            self.product_tree.delete(item)

    def on_tree_select(self, event):
        """当在商品列表中选择一项时，自动填充输入框"""
        selected = self.product_tree.selection()
        if not selected:
            return
        
        item = self.product_tree.item(selected[0])
        values = item['values']
        
        self.clear_entries()
        
        try:
            if len(values) >= 8:  # 确保包含出库金额值
                # 填充基本信息输入框
                self.id_entry.insert(0, values[0])
                self.name_entry.insert(0, values[1])
                self.price_entry.insert(0, values[2])
                self.colour_entry.insert(0, values[3])
                self.stock_entry.insert(0, values[4])
                self.date_entry.insert(0, values[5])
                # 更新剩余库存显示
                self.remaining_stock_label.config(text=str(values[6]))
                # 更新出库金额显示
                outbound_amount = float(values[7]) if values[7] else 0
                self.outbound_amount_label.config(text=f"¥{int(outbound_amount)}")  # 去掉小数位
                
                # 填充高级查询输入框
                self.search_id_entry.delete(0, tk.END)
                self.search_id_entry.insert(0, values[0])  # 产品型号
                
                self.search_name_entry.delete(0, tk.END)
                self.search_name_entry.insert(0, values[1])  # 产品甲方
                
        except Exception as e:
            messagebox.showerror("错误", f"填充输入框时出错：{str(e)}")

    def on_tree_focus_out(self, event):
        """当商品列表失去焦点时，清空输入框"""
        # 添加短暂延时，避免与选择事件冲突
        self.root.after(100, self.check_tree_selection)

    def check_tree_selection(self):
        """检查树形视图是否有选中项，如无则清空输入框"""
        if not self.product_tree.selection():
            self.clear_entries()
            if hasattr(self, 'selected_product_label'):
                self.selected_product_label.config(text="当前选中商品：未选择")

    def display_products(self):
        # 清空现有的treeview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 获取所有产品信息
        self.cursor.execute("""
            SELECT product_id, name, price, colour, stock, order_date 
            FROM products
        """)
        
        # 显示数据
        for row in self.cursor.fetchall():
            self.tree.insert("", "end", values=row)

    def setup_treeview(self):
        # 创建Treeview
        columns = ("product_id", "name", "price", "colour", "stock", "order_date")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        
        # 设置列标题
        self.tree.heading("product_id", text="产品ID")
        self.tree.heading("name", text="产品名称")
        self.tree.heading("price", text="价格")
        self.tree.heading("colour", text="颜色")
        self.tree.heading("stock", text="库存")
        self.tree.heading("order_date", text="订单日期")
        
        # ... existing code ...

    def enable_entries(self):
        """启用所有输入框"""
        self.id_entry.config(state='normal')
        self.name_entry.config(state='normal')
        self.price_entry.config(state='normal')
        self.colour_entry.config(state='normal')
        self.stock_entry.config(state='normal')
        self.date_entry.config(state='normal')


    def disable_entries(self):
        """禁用所有输入框"""
        self.id_entry.config(state='disabled')
        self.name_entry.config(state='disabled')
        self.price_entry.config(state='disabled')
        self.colour_entry.config(state='disabled')
        self.stock_entry.config(state='disabled')
        self.date_entry.config(state='disabled')

    def generate_current_date(self):
        """生成当前日期并填入日期输入框"""
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, current_date)

    def create_search_page(self):
        """创建查询页面"""
        search_frame = ttk.LabelFrame(self.notebook, text="查询产品")
        search_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # 搜索条件框架
        search_input_frame = ttk.Frame(search_frame)
        search_input_frame.pack(fill='x', padx=5, pady=5)
        
        # 创建树形视图
        self.search_tree = ttk.Treeview(search_frame, columns=("ID", "Name", "Price", "Colour", "Stock", "Date"), show="headings")
        self.search_tree.heading("ID", text="产品型号")
        self.search_tree.heading("Name", text="产品甲方")
        self.search_tree.heading("Price", text="产品价格")
        self.search_tree.heading("Colour", text="产品颜色")
        self.search_tree.heading("Stock", text="订单数量")
        self.search_tree.heading("Date", text="订单日期")
        
        # 设置列宽
        self.search_tree.column("ID", width=100)
        self.search_tree.column("Name", width=150)
        self.search_tree.column("Price", width=100)
        self.search_tree.column("Colour", width=100)
        self.search_tree.column("Stock", width=100)
        self.search_tree.column("Date", width=100)
        
        # ... 其他代码保持不变 ...

    def search_products(self):
        """搜索产品"""
        keyword = self.search_entry.get()
        
        # 清空现有结果
        for item in self.search_tree.get_children():
            self.search_tree.delete(item)
        
        try:
            # 使用明确的列名进行查询
            self.cursor.execute("""
                SELECT 
                    product_id,
                    name,
                    price,
                    colour,
                    stock,
                    order_date
                FROM products 
                WHERE product_id LIKE ? 
                   OR name LIKE ? 
                   OR colour LIKE ?
            """, (f"%{keyword}%", f"%{keyword}%", f"%{keyword}%"))
            
            results = self.cursor.fetchall()
            
            if not results:
                messagebox.showinfo("提示", "未找到匹配的产品")
                return
            
            # 显示结果并打印调试信息
            for row in results:
                print(f"Debug - Row data: {row}")  # 调试信息
                # 确保所有值都转换为字符串
                values = (
                    str(row[0]),  # product_id
                    str(row[1]),  # name
                    str(row[2]),  # price
                    str(row[3]),  # colour
                    str(row[4]),  # stock
                    str(row[5])   # order_date
                )
                self.search_tree.insert("", "end", values=values)
            
            # 打印列数检查
            print(f"Debug - Tree columns: {self.search_tree['columns']}")
            
        except Exception as e:
            messagebox.showerror("错误", f"搜索失败：{str(e)}")
            print(f"Debug - Error details: {str(e)}")  # 调试信息 

    def search_movements(self):
        """按条件搜索库存变动记录"""
        try:
            # 获取查询条件
            product_id = self.search_product_entry.get().strip()
            start_date = self.start_date_entry.get().strip()
            end_date = self.end_date_entry.get().strip()
            
            # 如果没有输入任何查询条件，显示提示信息
            if not any([product_id, start_date, end_date]):
                messagebox.showinfo("提示", "请输入至少一个查询条件")
                return
            
            # 清空现有显示
            for item in self.movement_tree.get_children():
                self.movement_tree.delete(item)
                
            try:
                cursor = self.conn.cursor()
                query = """
                    SELECT 
                        product_id,
                        movement_type,
                        quantity,
                        datetime,
                        remarks
                    FROM stock_movements
                    WHERE 1=1
                """
                params = []
                
                # 添加查询条件
                if product_id:
                    query += " AND product_id LIKE ?"
                    params.append(f"%{product_id}%")
                
                if start_date:
                    query += " AND remarks >= ?"
                    params.append(f"{start_date}")
                
                if end_date:
                    query += " AND remarks <= ?"
                    params.append(f"{end_date}")
                    
                query += " ORDER BY datetime DESC"
                
                cursor.execute(query, params)
                movements = cursor.fetchall()
                
                if not movements:
                    messagebox.showinfo("提示", "未找到符合条件的记录")
                    return
                
                for movement in movements:
                    product_id, movement_type, quantity, datetime_str, remarks = movement
                    
                    # 直接显示数量，不添加正负号
                    self.movement_tree.insert("", "end", values=(
                        product_id,
                        movement_type,
                        str(quantity),  # 直接显示数量
                        remarks,      # 出库日期
                        datetime_str  # 变动时间
                    ))
                    
            except sqlite3.Error as e:
                messagebox.showerror("错误", f"查询失败: {str(e)}")
                
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")
        
        # 更新状态栏
        self.update_stock_status_bar()
    
    def reset_stock_search(self):
        """重置搜索条件并显示所有记录"""
        # 清空搜索条件
        self.search_product_entry.delete(0, tk.END)
        self.start_date_entry.delete(0, tk.END)
        self.end_date_entry.delete(0, tk.END)
        
        # 清空输入框
        self.movement_id_entry.delete(0, tk.END)
        self.movement_type.set("")
        self.quantity_entry.delete(0, tk.END)
        self.remarks_entry.delete(0, tk.END)
        
        # 使用 view_stock_movements 来显示所有记录
        self.view_stock_movements()

    def reset_product_form(self):
        """重置产品表单所有输入框和选中状态"""
        # 清空所有输入框
        self.clear_entries()
        
        # 清空高级查询输入框
        if hasattr(self, 'search_id_entry'):
            self.search_id_entry.delete(0, tk.END)
        if hasattr(self, 'search_name_entry'):
            self.search_name_entry.delete(0, tk.END)
        if hasattr(self, 'search_start_date'):
            self.search_start_date.delete(0, tk.END)
        if hasattr(self, 'search_end_date'):
            self.search_end_date.delete(0, tk.END)
        
        # 清除树形视图的选中状态
        for item in self.product_tree.selection():
            self.product_tree.selection_remove(item)
        
        # 显示所有产品
        self.show_all_products()
       
    def advanced_movement_search(self):
        """库存变动高级查询"""
        # 创建高级查询对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("库存变动高级查询")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 禁用对话框的调整大小功能
        dialog.resizable(False, False)
        
        # 计算居中位置
        # 获取主窗口位置和大小
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        # 获取对话框大小
        dialog_width = 500
        dialog_height = 400
        
        # 计算对话框应该出现的位置
        x = root_x + (root_width - dialog_width) // 2
        y = root_y + (root_height - dialog_height) // 2
        
        # 设置对话框位置
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 查询条件框架
        search_frame = ttk.LabelFrame(dialog, text="查询条件")
        search_frame.pack(fill='x', padx=10, pady=5)
        
        # 产品信息
        product_frame = ttk.Frame(search_frame)
        product_frame.pack(fill='x', pady=5)
        ttk.Label(product_frame, text="产品型号:").pack(side='left', padx=5)
        product_entry = ttk.Entry(product_frame, width=15)
        product_entry.pack(side='left', padx=5)
        
        ttk.Label(product_frame, text="变动类型:").pack(side='left', padx=5)
        type_var = tk.StringVar()
        type_combo = ttk.Combobox(product_frame, textvariable=type_var, width=12)
        type_combo['values'] = ('全部', '入库', '出库')
        type_combo.set('全部')
        type_combo.pack(side='left', padx=5)
        
        # 数量范围
        quantity_frame = ttk.Frame(search_frame)
        quantity_frame.pack(fill='x', pady=5)
        ttk.Label(quantity_frame, text="数量范围:").pack(side='left', padx=5)
        quantity_min = ttk.Entry(quantity_frame, width=8)
        quantity_min.pack(side='left', padx=5)
        ttk.Label(quantity_frame, text="至").pack(side='left')
        quantity_max = ttk.Entry(quantity_frame, width=8)
        quantity_max.pack(side='left', padx=5)
        
        # 日期范围
        date_frame = ttk.Frame(search_frame)
        date_frame.pack(fill='x', pady=5)
        ttk.Label(date_frame, text="日期范围:").pack(side='left', padx=5)
        date_start = ttk.Entry(date_frame, width=12)
        date_start.pack(side='left', padx=5)
        ttk.Label(date_frame, text="至").pack(side='left')
        date_end = ttk.Entry(date_frame, width=12)
        date_end.pack(side='left', padx=5)
        
        # 快捷日期按钮
        quick_date_frame = ttk.Frame(search_frame)
        quick_date_frame.pack(fill='x', pady=5)
        
        def set_date_range(days):
            end_date = datetime.now()
            start_date = end_date - timedelta(days=days)
            date_start.delete(0, tk.END)
            date_start.insert(0, start_date.strftime("%Y-%m-%d"))
            date_end.delete(0, tk.END)
            date_end.insert(0, end_date.strftime("%Y-%m-%d"))
        
        ttk.Button(quick_date_frame, text="今天", 
                  command=lambda: set_date_range(0)).pack(side='left', padx=5)
        ttk.Button(quick_date_frame, text="最近7天", 
                  command=lambda: set_date_range(7)).pack(side='left', padx=5)
        ttk.Button(quick_date_frame, text="最近30天", 
                  command=lambda: set_date_range(30)).pack(side='left', padx=5)
        ttk.Button(quick_date_frame, text="最近90天", 
                  command=lambda: set_date_range(90)).pack(side='left', padx=5)
        
        def execute_search():
            try:
                # 构建查询条件
                conditions = []
                params = []
                
                # 产品型号
                if product_entry.get().strip():
                    conditions.append("product_id LIKE ?")
                    params.append(f"%{product_entry.get().strip()}%")
                
                # 变动类型
                if type_var.get() != "全部":
                    conditions.append("movement_type = ?")
                    params.append(type_var.get())
                
                # 数量范围
                if quantity_min.get().strip():
                    conditions.append("quantity >= ?")
                    params.append(int(quantity_min.get()))
                if quantity_max.get().strip():
                    conditions.append("quantity <= ?")
                    params.append(int(quantity_max.get()))
                
                # 日期范围
                if date_start.get().strip():
                    conditions.append("date(remarks) >= ?")
                    params.append(date_start.get())
                if date_end.get().strip():
                    conditions.append("date(remarks) <= ?")
                    params.append(date_end.get())
                
                # 构建SQL查询
                query = """
                    SELECT product_id, movement_type, quantity, remarks, datetime, notes
                    FROM stock_movements
                """
                if conditions:
                    query += " WHERE " + " AND ".join(conditions)
                query += " ORDER BY datetime DESC"
                
                # 执行查询
                self.cursor.execute(query, params)
                results = self.cursor.fetchall()
                
                # 清空并显示结果
                self.clear_movement_tree()
                for result in results:
                    self.movement_tree.insert("", "end", values=result)
                
                # 更新状态栏
                self.update_stock_status_bar()
                
                # 关闭对话框
                dialog.destroy()
                
                # 显示结果统计
                messagebox.showinfo("查询完成", f"共找到 {len(results)} 条记录")
                
            except Exception as e:
                messagebox.showerror("错误", f"查询失败：{str(e)}")
        
        # 按钮框架
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="执行查询", 
                  command=execute_search).pack(side='left', padx=5)
        ttk.Button(button_frame, text="取消", 
                  command=dialog.destroy).pack(side='left', padx=5)

    def update_movement(self):
        """修改选中的库存变动记录"""
        selected = self.movement_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要修改的记录！")
            return
        
        try:
            # 获取原始记录信息
            item = self.movement_tree.item(selected[0])
            old_values = item['values']
            
            # 打印调试信息
            print(f"Debug - Selected record values: {old_values}")
            
            # 创建修改记录对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("修改库存变动记录")
            dialog.geometry("400x400")  # 增加窗口高度以容纳备注
            dialog.transient(self.root)
            dialog.grab_set()
            
            # 禁用对话框的调整大小功能
            dialog.resizable(False, False)
            
            # 计算居中位置
            root_x = self.root.winfo_x()
            root_y = self.root.winfo_y()
            root_width = self.root.winfo_width()
            root_height = self.root.winfo_height()
            
            dialog_width = 400
            dialog_height = 400  # 增加高度
            
            x = root_x + (root_width - dialog_width) // 2
            y = root_y + (root_height - dialog_height) // 2
            
            dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
            
            # 创建输入框架
            input_frame = ttk.LabelFrame(dialog, text="修改信息", padding=10)
            input_frame.pack(fill='x', padx=10, pady=5)
            
            # 显示产品型号（只读）
            ttk.Label(input_frame, text="产品型号:").grid(row=0, column=0, padx=5, pady=5)
            product_id_label = ttk.Label(input_frame, text=old_values[0])
            product_id_label.grid(row=0, column=1, padx=5, pady=5)
            
            # 变动类型选择
            ttk.Label(input_frame, text="变动类型:").grid(row=1, column=0, padx=5, pady=5)
            type_var = tk.StringVar(value=old_values[1])
            type_combo = ttk.Combobox(input_frame, textvariable=type_var, width=15)
            type_combo['values'] = ('入库', '出库')
            type_combo['state'] = 'readonly'
            type_combo.grid(row=1, column=1, padx=5, pady=5)
            
            # 变动数量输入
            ttk.Label(input_frame, text="变动数量:").grid(row=2, column=0, padx=5, pady=5)
            quantity_entry = ttk.Entry(input_frame, width=18)
            quantity_entry.insert(0, old_values[2])
            quantity_entry.grid(row=2, column=1, padx=5, pady=5)
            
            # 库存日期输入
            ttk.Label(input_frame, text="库存日期:").grid(row=3, column=0, padx=5, pady=5)
            date_entry = ttk.Entry(input_frame, width=18)
            date_entry.insert(0, old_values[3])
            date_entry.grid(row=3, column=1, padx=5, pady=5)
            
            # 备注说明输入
            ttk.Label(input_frame, text="备注说明:").grid(row=4, column=0, padx=5, pady=5)
            notes_text = tk.Text(input_frame, width=30, height=4)
            if len(old_values) > 5 and old_values[5]:  # 如果有备注
                notes_text.insert("1.0", old_values[5])
            notes_text.grid(row=4, column=1, padx=5, pady=5)
            
            def save_changes():
                try:
                    # 验证输入
                    quantity_str = quantity_entry.get().strip()
                    if not quantity_str:
                        messagebox.showwarning("警告", "请输入变动数量！")
                        return
                    try:
                        quantity = int(quantity_str)
                        if quantity <= 0:
                            messagebox.showwarning("警告", "变动数量必须大于0！")
                            return
                    except ValueError:
                        messagebox.showwarning("警告", "变动数量必须是整数！")
                        return
                    
                    date_str = date_entry.get().strip()
                    if not date_str:
                        messagebox.showwarning("警告", "请输入库存日期！")
                        return
                    
                    movement_type = type_var.get()
                    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    notes = notes_text.get("1.0", tk.END).strip()  # 获取备注内容
                    
                    # 使用所有可用字段来定位记录
                    self.cursor.execute("""
                        SELECT rowid, * FROM stock_movements 
                        WHERE product_id = ? 
                        AND movement_type = ? 
                        AND quantity = ? 
                        AND remarks = ?
                    """, (old_values[0], old_values[1], old_values[2], old_values[3]))
                    
                    row = self.cursor.fetchone()
                    if not row:
                        messagebox.showerror("错误", "找不到要修改的记录！")
                        return
                    
                    rowid = row[0]
                    
                    # 开始事务
                    self.conn.execute("BEGIN")
                    try:
                        # 更新库存变化
                        old_quantity_change = -old_values[2] if old_values[1] == "出库" else old_values[2]
                        new_quantity_change = -quantity if movement_type == "出库" else quantity
                        
                        # 更新产品库存
                        net_change = new_quantity_change - old_quantity_change
                        self.cursor.execute("""
                            UPDATE products 
                            SET stock = stock + ? 
                            WHERE product_id = ?
                        """, (net_change, old_values[0]))
                        
                        # 使用 rowid 更新变动记录，包括备注
                        self.cursor.execute("""
                            UPDATE stock_movements 
                            SET movement_type = ?, 
                                quantity = ?, 
                                remarks = ?, 
                                datetime = ?,
                                notes = ?
                            WHERE rowid = ?
                        """, (movement_type, quantity, date_str, current_time, notes, rowid))
                        
                        # 提交事务
                        self.conn.commit()
                        
                        messagebox.showinfo("成功", "记录已更新！")
                        dialog.destroy()
                        
                        # 刷新显示
                        self.view_stock_movements()
                        self.show_all_products()
                        
                    except Exception as e:
                        # 发生错误时回滚事务
                        self.conn.rollback()
                        raise e
                        
                except Exception as e:
                    messagebox.showerror("错误", f"更新失败：{str(e)}")
                    print(f"Debug - Error details: {str(e)}")
            
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=20)
            
            ttk.Button(button_frame, text="保存", 
                      command=save_changes).pack(side='left', padx=10)
            ttk.Button(button_frame, text="取消", 
                      command=dialog.destroy).pack(side='left', padx=10)
            
        except Exception as e:
            messagebox.showerror("错误", f"打开修改窗口失败：{str(e)}")

    def delete_movement(self):
        """删除选中的库存变动记录"""
        selected = self.movement_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要删除的记录！")
            return
        
        if messagebox.askyesno("确认", "确定要删除选中的记录吗？\n注意：这将同时更新相关产品的库存！"):
            try:
                # 开始事务
                self.conn.execute("BEGIN")
                
                deleted_records = []  # 用于存储成功删除的记录信息
                
                for item_id in selected:
                    # 获取记录信息
                    item = self.movement_tree.item(item_id)
                    values = item['values']
                    
                    # 打印调试信息
                    print(f"Debug - 尝试删除记录: {values}")
                    
                    product_id = values[0]
                    movement_type = values[1]
                    quantity = int(values[2])
                    remarks = values[3]
                    datetime_str = values[4]
                    
                    # 首先检查记录是否存在
                    self.cursor.execute("""
                        SELECT COUNT(*) FROM stock_movements 
                        WHERE product_id = ? 
                        AND movement_type = ? 
                        AND quantity = ?
                        AND remarks = ?
                    """, (product_id, movement_type, quantity, remarks))
                    
                    count = self.cursor.fetchone()[0]
                    print(f"Debug - 找到匹配记录数: {count}")
                    
                    if count > 0:
                        # 更新产品库存
                        quantity_change = quantity if movement_type == "出库" else -quantity
                        
                        # 更新产品库存
                        self.cursor.execute("""
                            UPDATE products 
                            SET stock = stock + ? 
                            WHERE product_id = ?
                        """, (quantity_change, product_id))
                        
                        # 使用更宽松的条件删除记录
                        self.cursor.execute("""
                            DELETE FROM stock_movements 
                            WHERE product_id = ? 
                            AND movement_type = ? 
                            AND quantity = ?
                            AND remarks = ?
                        """, (product_id, movement_type, quantity, remarks))
                        
                        deleted_count = self.cursor.rowcount
                        print(f"Debug - 删除的记录数: {deleted_count}")
                        
                        if deleted_count > 0:
                            deleted_records.append(item_id)
                            # 从树形视图中删除
                            self.movement_tree.delete(item_id)
                    
                # 提交事务
                self.conn.commit()
                
                if deleted_records:
                    messagebox.showinfo("成功", f"已删除 {len(deleted_records)} 条记录！")
                    
                    # 清空输入框
                    self.clear_movement_entries()
                    
                    # 重新加载数据
                    self.show_all_products()  # 刷新产品列表
                    self.view_all_movements()  # 刷新库存变动记录
                    
                    # 更新状态栏
                    self.update_stock_status_bar()
                else:
                    print("Debug - 要删除的记录信息:")
                    for item_id in selected:
                        values = self.movement_tree.item(item_id)['values']
                        print(f"记录值: {values}")
                    messagebox.showwarning("警告", "没有记录被删除！可能是因为数据库中找不到匹配的记录。")
                
            except Exception as e:
                # 发生错误时回滚事务
                self.conn.rollback()
                messagebox.showerror("错误", f"删除失败：{str(e)}")
                print(f"Debug - Error details: {str(e)}")
    
    def on_stock_tree_select(self, event):
        """当选择库存明细表的条目时,自动填充输入框"""
        selected = self.stock_tree.selection()
        if not selected:
            return
        
        # 获取选中项的值
        item = self.stock_tree.item(selected[0])
        values = item['values']
        
        # 清空输入框
        self.clear_movement_entries()
        
        # 填充输入框
        if len(values) >= 3:  # 确保有足够的值
            self.movement_id_entry.insert(0, values[0])     # 产品型号
            self.quantity_entry.insert(0, values[2])        # 当前库存作为默认数量
            current_date = datetime.now().strftime("%Y-%m-%d")
            self.remarks_entry.insert(0, current_date)      # 当前日期作为默认出库日期
       
    def on_product_tree_select_for_stock(self, event):
        """当在库存变动页面选择产品记录时,自动填充产品型号"""
        selected = self.product_tree.selection()
        if not selected:
            return
        
        # 获取选中项的值
        item = self.product_tree.item(selected[0])
        values = item['values']
        
        if values:
            # 清空产品型号输入框
            self.movement_id_entry.delete(0, tk.END)
            # 填充产品型号
            self.movement_id_entry.insert(0, values[0])  # values[0] 是产品型号
            
            # 可选：自动填充当前日期
            current_date = datetime.now().strftime("%Y-%m-%d")
            self.remarks_entry.delete(0, tk.END)
            self.remarks_entry.insert(0, current_date)
       
    def on_movement_tree_select(self, event):
        """当选择变动记录时,自动填充输入框"""
        selected = self.movement_tree.selection()
        if not selected:
            return
        
        # 获取选中项的值
        item = self.movement_tree.item(selected[0])
        values = item['values']
        
        if not values:
            return
        
        # print(f"Debug - Selected values: {values}")  # 添加调试输出
        
        # 清空相关输入框
        self.movement_id_entry.delete(0, tk.END)
        self.movement_type.set("")
        self.quantity_entry.delete(0, tk.END)
        self.remarks_entry.delete(0, tk.END)
        self.notes_text.delete("1.0", tk.END)
        
        # 填充输入框
        try:
            self.movement_id_entry.insert(0, values[0])     # 产品型号
            self.movement_type.set(values[1])               # 变动类型
            self.quantity_entry.insert(0, values[2])        # 变动数量
            self.remarks_entry.insert(0, values[3])         # 库存日期
            if len(values) > 5 and values[5]:              # 如果有备注说明
                self.notes_text.insert("1.0", values[5])    # 填充备注说明
                
            # 在查询框中填充产品型号
            if hasattr(self, 'search_product_entry'):
                self.search_product_entry.delete(0, tk.END)
                self.search_product_entry.insert(0, values[0])
                
        except Exception as e:
            print(f"Debug - Error filling entries: {str(e)}")  # 添加错误调试输出
    
    def export_current_view(self):
        """导出当前视图中显示的记录"""
        try:
            # 获取当前树形视图中显示的所有记录
            data = []
            for item in self.movement_tree.get_children():
                values = self.movement_tree.item(item)['values']
                data.append({
                    '商品编号': values[0],
                    '变动类型': values[1],
                    '数量': values[2],
                    '库存日期': values[3],  # 修改这里的"出库日期"为"库存日期"
                    '变动时间': values[4]
                })
            
            if not data:
                messagebox.showinfo("提示", "当前没有可导出的数据！")
                return
            
            # 创建DataFrame
            import pandas as pd
            df = pd.DataFrame(data)
            
            # 生成默认文件名
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f'当前出入库明细表_{current_time}.xlsx'
            
            # 打开文件选择对话框
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if filename:  # 如果用户选择了保存路径
                df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"报表已导出至：{filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")
       
    def setup_custom_styles(self):
        """设置自定义样式"""
        style = ttk.Style()
        
        # 添加输入框样式
        style.configure('Entry.TEntry',
            padding=5,
            font=('Microsoft YaHei', 9)
        )
        
        # 添加标签样式
        style.configure('Label.TLabel',
            font=('Microsoft YaHei', 9),
            padding=2
        )
        
        # 修改按钮样式定义，所有文字改为黑色
        button_styles = {
            'primary': ('#0d6efd', '#0a58ca', '#000000'),    # 深蓝底黑字
            'info': ('#0dcaf0', '#0aa1c0', '#000000'),       # 浅蓝底黑字
            'danger': ('#dc3545', '#bb2d3b', '#000000'),     # 红底黑字
            'warning': ('#ffc107', '#ffca2c', '#000000'),    # 黄底黑字
            'success': ('#198754', '#157347', '#000000'),    # 绿底黑字
            'secondary': ('#6c757d', '#5c636a', '#000000')   # 灰底黑字
        }
        
        # 为每种按钮类型设置样式，增加文字阴影效果
        for name, (normal, hover, fg_color) in button_styles.items():
            style.configure(
                f'{name}.TButton',
                background=normal,
                foreground=fg_color,
                padding=(10, 5),
                font=('Microsoft YaHei', 9, 'bold'),
                relief='raised'
            )
            style.map(
                f'{name}.TButton',
                background=[('active', hover)],
                foreground=[('active', fg_color)],
                relief=[('pressed', 'sunken')]
            )
        
        # 设置表格样式
        style.configure(
            "Treeview",
            background="#ffffff",
            foreground="#333333",
            rowheight=25,
            fieldbackground="#ffffff"
        )
        style.configure(
            "Treeview.Heading",
            background="#f8f9fa",
            foreground="#333333",
            padding=(5, 5),
            font=('Microsoft YaHei', 9, 'bold')
        )
        
        # 设置选中行的样式
        style.map(
            "Treeview",
            background=[('selected', '#007bff')],
            foreground=[('selected', 'white')]
        )
       
    def reset_movement_form(self):
        """重置库存变动表单和查询条件"""
        try:
            # 清空变动登记区域
            self.movement_id_entry.delete(0, tk.END)
            self.movement_type.set("")
            self.quantity_entry.delete(0, tk.END)
            self.remarks_entry.delete(0, tk.END)
            self.notes_text.delete("1.0", tk.END)
            
            # 清空记录查询区域
            self.search_product_entry.delete(0, tk.END)
            self.start_date_entry.delete(0, tk.END)
            self.end_date_entry.delete(0, tk.END)
            
            # 重置树形视图显示
            self.clear_movement_tree()
            
            # 默认填入今天日期
            self.remarks_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
            
            # 更新状态栏
            self.update_stock_status_bar()
            
        except Exception as e:
            messagebox.showerror("错误", f"重置失败：{str(e)}")

    def view_current_movements(self):
        """只显示当前选中产品的库存变动记录"""
        try:
            self.clear_movement_tree()
            product_id = self.movement_id_entry.get()
            
            if not product_id:
                messagebox.showwarning("警告", "请先选择或输入产品型号！")
                return
            
            self.cursor.execute("""
                SELECT product_id, movement_type, quantity, remarks, datetime
                FROM stock_movements
                WHERE product_id=?
                ORDER BY datetime DESC
            """, (product_id,))
            
            movements = self.cursor.fetchall()
            if not movements:
                messagebox.showinfo("提示", "该产品没有库存变动记录")
                return
            
            for movement in movements:
                self.movement_tree.insert("", "end", values=movement)
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")
        
        # 更新状态栏
        self.update_stock_status_bar()

    def view_all_movements(self):
        """显示所有产品的库存变动记录，默认显示最近7天"""
        try:
            self.clear_movement_tree()
            
            self.cursor.execute("""
                SELECT product_id, movement_type, quantity, remarks, datetime, notes
                FROM stock_movements
                WHERE date(datetime) >= date('now', '-7 day')
                ORDER BY datetime DESC
            """)
            
            movements = self.cursor.fetchall()
            if not movements:
                return
            
            for movement in movements:
                product_id, movement_type, quantity, stock_date, datetime_str, notes = movement
                # 分离日期和时间
                datetime_obj = datetime.strptime(datetime_str, "%Y-%m-%d %H:%M:%S")
                formatted_time = datetime_obj.strftime("%Y-%m-%d %H:%M")
                
                self.movement_tree.insert("", "end", values=(
                    product_id,
                    movement_type,
                    str(quantity),
                    stock_date,  # 库存日期(只有日期)
                    formatted_time,  # 变动时间(包含时分)
                    notes or ""
                ))
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")
        
        # 更新状态栏
        self.update_stock_status_bar()

    def create_role_switch_button(self):
        """创建角色切换按钮"""
        self.role_button = ttk.Button(
            self.root,
            text="普通用户",
            command=self.switch_role,
            style='secondary.TButton'
        )
        self.role_button.pack(anchor='ne', padx=10, pady=5)
        
    def switch_role(self):
        """切换用户角色"""
        if not self.is_admin:
            # 创建密码输入对话框
            password_dialog = tk.Toplevel(self.root)
            password_dialog.title("管理员验证")
            password_dialog.geometry("300x150")
            password_dialog.transient(self.root)
            
            # 禁用对话框的调整大小功能
            password_dialog.resizable(False, False)
            
            # 计算居中位置
            root_x = self.root.winfo_x()
            root_y = self.root.winfo_y()
            root_width = self.root.winfo_width()
            root_height = self.root.winfo_height()
            
            dialog_width = 300
            dialog_height = 150
            
            x = root_x + (root_width - dialog_width) // 2
            y = root_y + (root_height - dialog_height) // 2
            
            # 设置对话框位置
            password_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
            
            # 设置对话框置顶
            password_dialog.grab_set()
            
            ttk.Label(password_dialog, text="请输入管理员密码:").pack(pady=10)
            password_entry = ttk.Entry(password_dialog, show="*")
            password_entry.pack(pady=5)
            
            def verify_password():
                if password_entry.get() == self.admin_password:
                    self.is_admin = True
                    self.role_button.configure(text="管理员", style='warning.TButton')
                    self.update_button_states()
                    password_dialog.destroy()
                    messagebox.showinfo("成功", "已切换至管理员模式")
                else:
                    messagebox.showerror("错误", "密码错误")
                    
            ttk.Button(password_dialog, text="确认", command=verify_password).pack(pady=10)
            
            # 设置回车键触发确认按钮
            password_entry.bind('<Return>', lambda e: verify_password())
            
            # 设置焦点到密码输入框
            password_entry.focus()
            
        else:
            # 切换回普通用户
            self.is_admin = False
            self.role_button.configure(text="普通用户", style='secondary.TButton')
            self.update_button_states()
            messagebox.showinfo("成功", "已切换至普通用户模式")
            
    def update_button_states(self):
        """更新按钮状态"""
        # 直接使用存储的按钮引用更新状态
        for btn in self.admin_buttons:
            btn.configure(state='normal' if self.is_admin else 'disabled')

    def advanced_search(self):
        """执行高级查询"""
        try:
            # 获取所有查询条件
            product_id = self.search_id_entry.get().strip()
            name = self.search_name_entry.get().strip()
            start_date = self.search_start_date.get().strip()
            end_date = self.search_end_date.get().strip()
            
            # 检查是否至少有一个查询条件
            if not any([product_id, name, start_date, end_date]):
                messagebox.showwarning("警告", "请至少输入一个查询条件！")
                return
            
            # 构建查询条件
            conditions = []
            params = []
            
            if product_id:
                conditions.append("p.product_id LIKE ?")
                params.append(f"%{product_id}%")
            
            if name:
                conditions.append("p.name LIKE ?")
                params.append(f"%{name}%")
            
            # 处理开始日期，确保格式统一
            if start_date:
                try:
                    formatted_start = datetime.strptime(start_date, "%Y-%m-%d").strftime("%Y-%m-%d")
                    conditions.append("date(p.order_date) >= date(?)")
                    params.append(formatted_start)
                except ValueError:
                    messagebox.showerror("错误", "开始日期格式无效，请使用 YYYY-MM-DD 格式")
                    return
            
            # 处理结束日期，确保格式统一
            if end_date:
                try:
                    formatted_end = datetime.strptime(end_date, "%Y-%m-%d").strftime("%Y-%m-%d")
                    conditions.append("date(p.order_date) <= date(?)")
                    params.append(formatted_end)
                except ValueError:
                    messagebox.showerror("错误", "结束日期格式无效，请使用 YYYY-MM-DD 格式")
                    return
            
            # 构建SQL查询，包含剩余库存计算
            query = """
                SELECT p.product_id, p.name, p.price, p.colour, p.stock,
                       p.order_date,
                       (p.stock - COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           WHEN m.movement_type = '入库' THEN -m.quantity 
                           ELSE 0 
                       END), 0)) as remaining_stock,
                       (COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           ELSE 0 
                       END), 0) * p.price) as outbound_amount
                    FROM products p
                    LEFT JOIN stock_movements m ON p.product_id = m.product_id
            """
            
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            
            query += " GROUP BY p.product_id, p.name, p.price, p.colour, p.stock, p.order_date"
            
            # 执行查询
            self.cursor.execute(query, params)
            results = self.cursor.fetchall()
            
            # 清空并显示结果
            self.clear_tree()
            for result in results:
                self.product_tree.insert("", "end", values=result)
            
            # 更新状态栏，显示当前查询结果数量
            self.update_status_bar(len(results))
            
            if not results:
                messagebox.showinfo("提示", "未找到符合条件的产品")
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败：{str(e)}")

    def export_search_results(self):
        """导出当前查询结果"""
        try:
            # 获取当前显示的所有记录
            items = self.product_tree.get_children()
            if not items:
                # 检查是否有查询条件
                product_id = self.search_id_entry.get().strip()
                name = self.search_name_entry.get().strip()
                start_date = self.search_start_date.get().strip()
                end_date = self.search_end_date.get().strip()
                
                if not any([product_id, name, start_date, end_date]):
                    messagebox.showwarning("警告", "请先进行查询！至少需要输入一个查询条件。")
                else:
                    messagebox.showinfo("提示", "没有符合条件的数据可导出")
                return
            
            # 准备数据
            data = []
            for item in items:
                values = self.product_tree.item(item)['values']
                data.append({
                    '产品型号': values[0],
                    '产品甲方': values[1],
                    '产品价格': values[2],
                    '产品颜色': values[3],
                    '订单数量': values[4],
                    '订单日期': values[5],
                    '剩余库存': values[6] if len(values) > 6 else '',
                    '出库金额': values[7] if len(values) > 7 else ''
                })
            
            # 导出到Excel
            import pandas as pd
            df = pd.DataFrame(data)
            
            # 生成默认文件名
            current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"产品查询结果_{current_time}.xlsx"
            
            # 打开文件选择对话框
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_filename,
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if filename:  # 如果用户选择了保存路径
                df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"数据已导出至：{filename}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出失败：{str(e)}")

    def generate_order(self):
        """按订单时间排序显示产品,点击切换升序/降序"""
        try:
            # 清空当前显示
            self.clear_tree()
            
            # 根据排序状态决定排序方式
            order_direction = "DESC" if self.order_desc else "ASC"
            
            # 执行查询
            self.cursor.execute(f"""
                SELECT p.product_id, p.name, p.price, p.colour, p.stock,
                       p.order_date,
                       (p.stock - COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           WHEN m.movement_type = '入库' THEN -m.quantity 
                           ELSE 0 
                       END), 0)) as remaining_stock,
                       (COALESCE(SUM(CASE 
                           WHEN m.movement_type = '出库' THEN m.quantity 
                           ELSE 0 
                       END), 0) * p.price) as outbound_amount
                FROM products p
                LEFT JOIN stock_movements m ON p.product_id = m.product_id
                GROUP BY p.product_id, p.name, p.price, p.colour, p.stock, p.order_date
                ORDER BY date(p.order_date) {order_direction}
            """)
            
            # 显示排序后的结果
            results = self.cursor.fetchall()
            for result in results:
                self.product_tree.insert("", "end", values=result)
            
            # 切换排序状态
            self.order_desc = not self.order_desc
            
            # 更新按钮文本显示当前排序状态
            for child in self.product_frame.winfo_children():
                if isinstance(child, ttk.Frame):  # 查找按钮框架
                    for btn in child.winfo_children():
                        if isinstance(btn, ttk.Button) and btn['text'] in ["时间↓", "时间↑"]:
                            btn.configure(text="时间↓" if self.order_desc else "时间↑")
            
            if not results:
                messagebox.showinfo("提示", "没有产品数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"排序失败：{str(e)}")

    def sort_movements_by_time(self):
        """按时间排序显示库存变动记录,点击切换升序/降序"""
        try:
            self.clear_movement_tree()
            
            # 根据排序状态决定排序方式
            order_direction = "DESC" if self.movement_order_desc else "ASC"
            
            # 执行查询
            self.cursor.execute(f"""
                SELECT product_id, movement_type, quantity, remarks, datetime
                FROM stock_movements
                ORDER BY datetime {order_direction}
            """)
            
            movements = self.cursor.fetchall()
            if not movements:
                messagebox.showinfo("提示", "暂无库存变动记录")
                return
            
            for movement in movements:
                self.movement_tree.insert("", "end", values=movement)
            
            # 切换排序状态
            self.movement_order_desc = not self.movement_order_desc
            
            # 更新按钮文本显示当前排序状态
            for child in self.stock_frame.winfo_children():
                if isinstance(child, ttk.Frame):  # 查找按钮框架
                    for btn in child.winfo_children():
                        if isinstance(btn, ttk.Button) and btn['text'] in ["时间↓", "时间↑"]:
                            btn.configure(text="时间↓" if self.movement_order_desc else "时间↑")
            
        except Exception as e:
            messagebox.showerror("错误", f"排序失败：{str(e)}")
        
        # 更新状态栏
        self.update_stock_status_bar()

    # 添加快速入库方法
    def quick_stock_in(self):
        """快速入库功能"""
        # 获取选中的商品
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择要入库的商品！")
            return
        
        item = self.product_tree.item(selected[0])
        values = item['values']
        product_id = values[0]
        
        # 创建快速入库对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("快速入库")
        dialog.geometry("300x200")
        dialog.transient(self.root)  # 设置为主窗口的子窗口
        dialog.grab_set()  # 模态对话框
        dialog.resizable(False, False)  # 禁止调整大小
        
        # 计算居中位置
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        
        dialog_width = 300
        dialog_height = 200
        
        x = root_x + (root_width - dialog_width) // 2
        y = root_y + (root_height - dialog_height) // 2
        
        # 设置对话框位置
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
        
        # 显示商品信息
        info_frame = ttk.Frame(dialog)
        info_frame.pack(pady=10)
        
        ttk.Label(info_frame, text=f"商品型号: {product_id}").pack()
        ttk.Label(info_frame, text=f"当前库存: {values[6]}").pack()
        
        # 创建输入框架
        input_frame = ttk.Frame(dialog)
        input_frame.pack(pady=10)
        
        # 入库数量行
        quantity_frame = ttk.Frame(input_frame)
        quantity_frame.pack(fill='x', pady=5)
        ttk.Label(quantity_frame, text="入库数量:").pack(side='left', padx=5)
        quantity_entry = ttk.Entry(quantity_frame, width=15)
        quantity_entry.pack(side='left', padx=5)
        
        # 入库日期行
        date_frame = ttk.Frame(input_frame)
        date_frame.pack(fill='x', pady=5)
        ttk.Label(date_frame, text="入库日期:").pack(side='left', padx=5)
        date_entry = ttk.Entry(date_frame, width=15)
        date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))
        date_entry.pack(side='left', padx=5)
        
        def confirm_stock_in():
            try:
                quantity = int(quantity_entry.get())
                if quantity <= 0:
                    messagebox.showerror("错误", "入库数量必须大于0")
                    return
                
                # 记录入库
                self.cursor.execute("""
                    INSERT INTO stock_movements 
                    (product_id, movement_type, quantity, remarks, datetime, notes)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (product_id, "入库", quantity, 
                      date_entry.get(),  # 入库日期
                      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),  # 变动时间
                      self.notes_text.get("1.0", tk.END).strip()))  # 备注说明
                  
                self.conn.commit()
                dialog.destroy()
                
                # 刷新显示
                self.show_all_products()
                self.view_all_movements()
                messagebox.showinfo("成功", f"已完成{quantity}件商品的入库操作")
                
            except ValueError:
                messagebox.showerror("错误", "请输入有效的入库数量")
        
        # 确认按钮
        ttk.Button(dialog, text="确认入库", command=confirm_stock_in).pack(pady=10)
        
        # 绑定回车键
        quantity_entry.bind('<Return>', lambda e: confirm_stock_in())
        
        # 设置初始焦点
        quantity_entry.focus()

    def sort_treeview(self, col):
        """对表格按指定列排序"""
        try:
            # 获取所有数据
            data = [(self.product_tree.set(item, col), item) for item in self.product_tree.get_children('')]
            
            # 根据列的类型进行适当的转换
            if col in ["price", "stock", "remaining_stock", "outbound_amount"]:
                # 数值类型排序
                data = [(float(value.replace('¥', '')) if value and value != '¥0.00' else 0, item) 
                       for value, item in data]
            elif col == "order_date":
                # 日期类型排序
                from datetime import datetime
                data = [(datetime.strptime(value, "%Y-%m-%d") if value else datetime.min, item) 
                       for value, item in data]
            
            # 切换排序方向
            self.sort_states[col] = not self.sort_states[col]
            
            # 排序数据
            data.sort(reverse=self.sort_states[col])
            
            # 重新排列数据
            for idx, (_, item) in enumerate(data):
                self.product_tree.move(item, '', idx)
                
            # 更新表头显示排序方向
            for column in self.product_tree["columns"]:
                if column == col:
                    arrow = "↓" if self.sort_states[col] else "↑"
                    text = self.product_tree.heading(column)["text"]
                    if "↓" in text or "↑" in text:
                        text = text.rstrip("↓↑")
                    self.product_tree.heading(column, text=f"{text}{arrow}")
                else:
                    # 移除其他列的排序指示器
                    text = self.product_tree.heading(column)["text"]
                    if "↓" in text or "↑" in text:
                        text = text.rstrip("↓↑")
                    self.product_tree.heading(column, text=text)
                    
        except Exception as e:
            messagebox.showerror("错误", f"排序失败：{str(e)}")

    # 添加新的方法
    def download_template(self):
        """下载产品信息模板"""
        try:
            import pandas as pd
            
            # 创建模板数据
            template_data = {
                '产品型号': ['示例: P001'],
                '产品甲方': ['示例: 公司A'],
                '产品价格': ['示例: 100.00'],
                '产品颜色': ['示例: 红色'],
                '订单数量': ['示例: 50'],
                '订单日期': ['示例: 2024-03-21']
            }
            
            df = pd.DataFrame(template_data)
            
            # 打开文件选择对话框
            from tkinter import filedialog
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="产品信息模板.xlsx",
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if filename:
                df.to_excel(filename, index=False)
                messagebox.showinfo("成功", f"模板已保存至：{filename}")
                
        except Exception as e:
            messagebox.showerror("错误", f"模板下载失败：{str(e)}")

    def upload_products(self):
        """上传产品信息"""
        try:
            from tkinter import filedialog
            import pandas as pd
            
            # 打开文件选择对话框
            filename = filedialog.askopenfilename(
                filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not filename:
                return
                
            # 读取Excel文件
            df = pd.read_excel(filename)
            
            # 验证列名
            required_columns = ['产品型号', '产品甲方', '产品价格', '产品颜色', '订单数量', '订单日期']
            if not all(col in df.columns for col in required_columns):
                messagebox.showerror("错误", "Excel文件格式不正确，请使用下载的模板文件")
                return
                
            # 开始插入数据
            success_count = 0
            error_count = 0
            error_messages = []
            
            for index, row in df.iterrows():
                try:
                    # 验证数据
                    if pd.isna(row['产品型号']) or pd.isna(row['产品甲方']):
                        error_messages.append(f"第{index+2}行：产品型号和产品甲方不能为空")
                        error_count += 1
                        continue
                        
                    # 转换数据类型
                    price = float(row['产品价格'])
                    stock = int(row['订单数量'])
                    
                    # 插入数据
                    self.cursor.execute("""
                        INSERT INTO products (product_id, name, price, colour, stock, order_date)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (
                        str(row['产品型号']),
                        str(row['产品甲方']),
                        price,
                        str(row['产品颜色']),
                        stock,
                        str(row['订单日期'])
                    ))
                    
                    success_count += 1
                    
                except Exception as e:
                    error_count += 1
                    error_messages.append(f"第{index+2}行：{str(e)}")
            
            # 提交事务
            self.conn.commit()
            
            # 刷新显示
            self.show_all_products()
            
            # 显示结果
            result_message = f"上传完成\n成功：{success_count}条\n失败：{error_count}条"
            if error_messages:
                result_message += "\n\n错误详情：\n" + "\n".join(error_messages[:5])
                if len(error_messages) > 5:
                    result_message += "\n..."
            
            if error_count > 0:
                messagebox.showwarning("上传结果", result_message)
            else:
                messagebox.showinfo("上传结果", result_message)
                
        except Exception as e:
            messagebox.showerror("错误", f"上传失败：{str(e)}")

    # 添加新方法
    def open_stock_page(self):
        """打开库存变动页面并填充选中产品信息"""
        # 获取当前选中的产品
        selected = self.product_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请先选择产品！")
            return
            
        # 获取产品信息
        item = self.product_tree.item(selected[0])
        product_id = item['values'][0]
        
        # 切换到库存变动页面
        self.notebook.select(self.stock_frame)
        
        # 填充产品信息到库存变动登记
        self.movement_id_entry.delete(0, tk.END)
        self.movement_id_entry.insert(0, product_id)
        
        # 设置当前日期
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.remarks_entry.delete(0, tk.END)
        self.remarks_entry.insert(0, current_date)
        
        # 自动显示该产品的库存变动记录
        self.view_current_movements()

    def open_product_page(self):
        """打开产品管理页面并显示当前产品信息"""
        # 获取当前选中的库存记录
        selected = self.movement_tree.selection()
        if not selected:
            # 如果没有选中的记录，则使用变动登记框中的产品型号
            product_id = self.movement_id_entry.get().strip()
            if not product_id:
                messagebox.showwarning("警告", "请先选择一条记录或输入产品型号！")
                return
        else:
            # 获取选中记录的产品型号
            item = self.movement_tree.item(selected[0])
            product_id = item['values'][0]
        
        try:
            # 切换到产品管理页面
            self.notebook.select(self.product_frame)
            
            # 清空现有查询条件
            self.search_id_entry.delete(0, tk.END)
            self.search_name_entry.delete(0, tk.END)
            self.search_start_date.delete(0, tk.END)
            self.search_end_date.delete(0, tk.END)
            
            # 填充产品型号到查询条件
            self.search_id_entry.insert(0, product_id)
            
            # 执行查询
            self.advanced_search()
            
            # 确保选中查询到的产品
            for item in self.product_tree.get_children():
                if self.product_tree.item(item)['values'][0] == product_id:
                    self.product_tree.selection_set(item)
                    self.product_tree.see(item)  # 确保选中的项目可见
                    # 触发选择事件以填充产品信息
                    self.on_tree_select(None)
                    break
                    
        except Exception as e:
            messagebox.showerror("错误", f"跳转失败：{str(e)}")

    def update_status_bar(self, current_display_count=None):
        """更新状态栏信息"""
        try:
            # 获取产品统计信息
            self.cursor.execute("""
                SELECT 
                    COUNT(*) as total_products,
                    COUNT(CASE WHEN order_date >= date('now', '-30 day') THEN 1 END) as recent_products
                FROM products
            """)
            stats = self.cursor.fetchone()
            
            # 获取所有时间的剩余库存
            self.cursor.execute("""
                SELECT 
                    (SELECT COALESCE(SUM(
                        stock - COALESCE((
                            SELECT SUM(CASE 
                                WHEN movement_type = '出库' THEN quantity 
                                WHEN movement_type = '入库' THEN -quantity 
                                ELSE 0 
                            END)
                            FROM stock_movements sm 
                            WHERE sm.product_id = p.product_id
                        ), 0)
                    ), 0)
                    FROM products p) as total_remaining_stock
                FROM stock_movements
                LIMIT 1
            """)
            all_time_stats = self.cursor.fetchone()
            
            # 计算当前显示记录的出库金额
            current_outbound_amount = 0
            if hasattr(self, 'product_tree'):
                for item in self.product_tree.get_children():
                    values = self.product_tree.item(item)['values']
                    if len(values) >= 8:  # 确保有足够的列
                        outbound_amount = values[7]  # 假设出库金额在第8列
                        if outbound_amount:
                            current_outbound_amount += float(outbound_amount)
            
            total_remaining_stock = all_time_stats[0]     # 总剩余库存
            
            # 格式化状态栏文本
            status_text = (
                f"产品总数: {stats[0]} | "
                f"当前显示: {current_display_count or stats[1]} | "
                f"总剩余库存: {total_remaining_stock} | "
                f"当前出库金额: ¥{int(current_outbound_amount):,} | "  # 使用 int() 去掉小数位
                f"最近30天新增: {stats[1]} 个产品"
            )
            
            # 更新或创建状态栏
            if hasattr(self, 'status_label'):
                self.status_label.config(text=status_text)
            else:
                self.status_label = ttk.Label(
                    self.product_frame, 
                    text=status_text,
                    relief="sunken",
                    padding=(5, 2)
                )
                self.status_label.pack(side="bottom", fill="x", padx=5, pady=2)
                
        except Exception as e:
            print(f"更新状态栏失败：{str(e)}")

    def validate_date(self, P):
        """验证日期格式"""
        if P == "": return True  # 允许空值
        try:
            if len(P) <= 10:  # 只在输入完整日期时验证
                datetime.strptime(P, "%Y-%m-%d")
            return True
        except ValueError:
            return False

    def set_today_date(self):
        """设置今天日期"""
        self.remarks_entry.delete(0, tk.END)
        self.remarks_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

    def update_stock_status_bar(self):
        """更新库存变动状态栏信息"""
        try:
            # 获取当前显示的记录数
            current_records = len(self.movement_tree.get_children())
            
            # 计算当前显示记录的出入库数量和金额
            total_in = 0
            total_out = 0
            total_amount = 0
            
            # 获取今天的记录数
            today = datetime.now().strftime("%Y-%m-%d")
            today_records = 0
            
            for item in self.movement_tree.get_children():
                values = self.movement_tree.item(item)['values']
                if not values:  # 跳过空值
                    continue
                
                # 计算出入库数量
                if values[1] == "入库":
                    total_in += int(values[2])
                else:  # 出库
                    total_out += int(values[2])
                    # 计算出库金额
                    try:
                        self.cursor.execute("SELECT price FROM products WHERE product_id=?", (values[0],))
                        price_result = self.cursor.fetchone()
                        if price_result:  # 确保找到价格
                            total_amount += int(values[2]) * price_result[0]
                    except Exception as e:
                        print(f"计算金额时出错: {str(e)}")
                
                # 统计今天的记录
                if values[3] == today:
                    today_records += 1
            
            # 获取最近7天的记录数
            try:
                self.cursor.execute("""
                    SELECT COUNT(*) 
                    FROM stock_movements 
                    WHERE date(remarks) >= date('now', '-7 day')
                """)
                recent_result = self.cursor.fetchone()
                recent_records = recent_result[0] if recent_result else 0
            except Exception as e:
                print(f"获取最近记录数时出错: {str(e)}")
                recent_records = 0
            
            # 格式化状态栏文本
            status_text = (
                f"今日: {today_records}条  "
                f"本周: {recent_records}条  |  "
                f"入库: {total_in}  "
                f"出库: {total_out}  |  "
                f"出库金额: ¥{int(total_amount):,}  |  "
                f"当前显示: {current_records}条记录"
            )
            
            # 确保状态栏标签存在
            if hasattr(self, 'stock_status_label'):
                self.stock_status_label.config(text=status_text)
            else:
                # 如果状态栏标签不存在，创建它
                self.stock_status_label = ttk.Label(
                    self.stock_frame, 
                    text=status_text,
                    relief="sunken",
                    padding=(5, 2)
                )
                self.stock_status_label.pack(side="bottom", fill="x", padx=5, pady=2)
                
        except Exception as e:
            print(f"更新库存变动状态栏失败：{str(e)}")
            # 设置一个基本的状态文本
            basic_status = "状态栏更新失败"
            if hasattr(self, 'stock_status_label'):
                self.stock_status_label.config(text=basic_status)

    def sort_movement_tree(self, col):
        """对库存变动记录表格按指定列排序"""
        try:
            # 获取所有数据
            data = [(self.movement_tree.set(item, col), item) for item in self.movement_tree.get_children('')]
            
            # 根据列的类型进行适当的转换
            if col == "quantity":
                # 数值类型排序
                data = [(int(value) if value else 0, item) for value, item in data]
            elif col in ["stock_date", "datetime"]:
                # 日期类型排序
                data = [(datetime.strptime(value, "%Y-%m-%d" if col == "stock_date" else "%Y-%m-%d %H:%M") 
                        if value else datetime.min, item) for value, item in data]
            
            # 切换排序方向
            self.movement_sort_states[col] = not self.movement_sort_states[col]
            
            # 排序数据
            data.sort(reverse=self.movement_sort_states[col])
            
            # 重新排列数据
            for idx, (_, item) in enumerate(data):
                self.movement_tree.move(item, '', idx)
                
            # 更新表头显示排序方向
            for column in self.movement_tree["columns"]:
                if column == col:
                    arrow = "↓" if self.movement_sort_states[col] else "↑"
                    text = self.movement_tree.heading(column)["text"]
                    if "↓" in text or "↑" in text:
                        text = text.rstrip("↓↑")
                    self.movement_tree.heading(column, text=f"{text}{arrow}")
                else:
                    # 移除其他列的排序指示器
                    text = self.movement_tree.heading(column)["text"]
                    if "↓" in text or "↑" in text:
                        text = text.rstrip("↓↑")
                    self.movement_tree.heading(column, text=text)
                
        except Exception as e:
            messagebox.showerror("错误", f"排序失败：{str(e)}")

if __name__ == "__main__":
        app = ProductManagement()
        app.root.mainloop() 