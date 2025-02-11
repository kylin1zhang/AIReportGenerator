import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
from google import genai
import json

class MonthlyReportGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("月报生成器")
        self.root.geometry("800x600")
        
        # 从配置文件读取API key
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
                self.api_key = config['GEMINI_API_KEY']
        except:
            print("请先配置API key")
        
        # 创建主界面
        self.create_gui()
        
    def create_gui(self):
        # 创建标签页
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)
        
        # 数据导入页
        self.import_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.import_frame, text='导入历史数据')
        
        # 生成页
        self.generate_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.generate_frame, text='生成新月报')
        
        self.setup_import_page()
        self.setup_generate_page()
    
    def setup_import_page(self):
        # 导入按钮
        ttk.Button(self.import_frame, text="导入Excel文件", 
                  command=self.import_excel).pack(pady=10)
        
        # 预览区域
        self.preview_text = tk.Text(self.import_frame, height=20)
        self.preview_text.pack(fill='both', expand=True, padx=5, pady=5)
    
    def setup_generate_page(self):
        # 添加月份选择
        month_frame = ttk.Frame(self.generate_frame)
        month_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(month_frame, text="选择月份：").pack(side='left')
        self.month_var = tk.StringVar(value="Dec-24")
        self.month_combo = ttk.Combobox(month_frame, 
                                      textvariable=self.month_var,
                                      values=["Sep-24", "Oct-24", "Nov-24", "Dec-24"])
        self.month_combo.pack(side='left', padx=5)
        
        # 添加Development Area选择
        area_frame = ttk.Frame(self.generate_frame)
        area_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(area_frame, text="Development Area：").pack(side='left')
        self.area_var = tk.StringVar()
        self.area_combo = ttk.Combobox(area_frame, textvariable=self.area_var)
        self.area_combo.pack(side='left', padx=5)
        
        # 添加Item选择
        item_frame = ttk.Frame(self.generate_frame)
        item_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(item_frame, text="选择Item：").pack(side='left')
        self.item_var = tk.StringVar()
        self.item_combo = ttk.Combobox(item_frame, textvariable=self.item_var, width=50)
        self.item_combo.pack(side='left', padx=5)
        
        # 添加输入区域
        input_frame = ttk.LabelFrame(self.generate_frame, text="填写本月进展")
        input_frame.pack(fill='x', padx=5, pady=5)
        
        # Current Progress输入
        ttk.Label(input_frame, text="Current Progress:").pack(pady=2)
        self.progress_text = tk.Text(input_frame, height=5)
        self.progress_text.pack(fill='x', padx=5, pady=5)
        
        # Acting Plan输入
        ttk.Label(input_frame, text="Acting Plan:").pack(pady=2)
        self.plan_text = tk.Text(input_frame, height=5)
        self.plan_text.pack(fill='x', padx=5, pady=5)
        
        # 生成按钮
        ttk.Button(self.generate_frame, text="生成月报", 
                  command=self.generate_report).pack(pady=10)
        
        # 结果显示区
        result_frame = ttk.LabelFrame(self.generate_frame, text="生成结果")
        result_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.result_text = tk.Text(result_frame)
        self.result_text.pack(fill='both', expand=True, padx=5, pady=5)
    
    def import_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                # 读取Excel文件
                self.df = pd.read_excel(file_path)
                
                # 找到第一个Item所在的行
                start_row = None
                for idx, row in self.df.iterrows():
                    for col in self.df.columns:
                        val = str(row[col])
                        if 'ltem1:' in val or 'Item1:' in val:
                            start_row = idx
                            break
                    if start_row is not None:
                        break
                
                if start_row is not None:
                    # 从找到的Item1开始读取数据
                    self.df = pd.read_excel(file_path, skiprows=start_row)
                    
                    # 设置列名
                    self.df.columns = [
                        'Development Area', 'Details', 'Target Date',
                        'Sep-24 Current Progress', 'Sep-24 Acting Plan', 'Sep-24 Manager Comments', 'Sep-24 RAG',
                        'Oct-24 Current Progress', 'Oct-24 Acting Plan', 'Oct-24 Manager Comments', 'Oct-24 RAG',
                        'Nov-24 Current Progress', 'Nov-24 Acting Plan', 'Nov-24 Manager Comments', 'Nov-24 RAG',
                        'Dec-24 Current Progress', 'Dec-24 Acting Plan', 'Dec-24 Manager Comments', 'Dec-24 RAG'
                    ]
                    
                    # 处理Development Area
                    for idx, row in self.df.iterrows():
                        details = str(row['Details']).strip()
                        # 根据Item编号设置Development Area
                        if any(f'ltem{i}:' in details or f'Item{i}:' in details for i in range(1, 6)):
                            self.df.loc[idx, 'Development Area'] = 'Communication'
                        elif any(f'ltem{i}:' in details or f'Item{i}:' in details for i in range(6, 11)):
                            self.df.loc[idx, 'Development Area'] = 'Tech skillset'
                        elif any(f'ltem{i}:' in details or f'Item{i}:' in details for i in range(11, 17)):
                            self.df.loc[idx, 'Development Area'] = 'Project'
                        elif 'ltem17:' in details or 'Item17:' in details:  # 直接使用 or 运算符
                            self.df.loc[idx, 'Development Area'] = 'EE Principles'
                    
                    # 填充空的Development Area
                    self.df['Development Area'] = self.df['Development Area'].ffill()
                    
                    # 打印调试信息
                    print("\n=== 导入的数据 ===")
                    print("数据形状:", self.df.shape)
                    print("\n前几行数据:")
                    for col in ['Development Area', 'Details', 'Sep-24 Current Progress', 'Sep-24 Acting Plan']:
                        print(f"\n{col}:")
                        print(self.df[col].head())
                    print("\n列名:", self.df.columns.tolist())
                    print("=============\n")
                    
                    # 更新界面
                    areas = ["Communication", "Tech skillset", "Project", "EE Principles"]
                    self.area_combo['values'] = areas
                    self.area_combo.bind('<<ComboboxSelected>>', self.on_area_selected)
                    
                    if self.area_var.get():
                        self.on_area_selected()
                    
                    self.preview_text.delete(1.0, tk.END)
                    self.preview_text.insert(tk.END, "数据导入成功！\n\n")
                    self.preview_text.insert(tk.END, str(self.df))
                else:
                    raise Exception("找不到Item1")
                
            except Exception as e:
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(tk.END, f"导入错误：{str(e)}")
                print(f"详细错误：{str(e)}")
                import traceback
                print(traceback.format_exc())
    
    def on_area_selected(self, event=None):
        """当选择Development Area时更新Items列表"""
        selected_area = self.area_var.get()
        if selected_area and hasattr(self, 'df'):
            try:
                print(f"\n=== 更新Items ===")
                print(f"选择的Area: {selected_area}")
                
                # 过滤该区域的数据
                area_data = self.df[self.df['Development Area'] == selected_area]
                print(f"\n该区域的数据:\n{area_data[['Development Area', 'Details']]}")
                
                # 收集所有Items
                items = []
                current_item = None
                
                for _, row in area_data.iterrows():
                    details = str(row.get('Details', '')).strip()
                    if pd.notna(details) and details:  # 检查是否为空
                        # 检查是否以Item或ltem开头（不区分大小写）
                        if any(details.lower().startswith(prefix) for prefix in ['item', 'ltem', 'item ', 'ltem ']):
                            if current_item:
                                items.append(current_item)
                            current_item = details
                        else:
                            # 如果当前行不是新的Item，则添加到当前Item
                            if current_item:
                                current_item = f"{current_item}\n{details}"
                
                # 添加最后一个Item
                if current_item:
                    items.append(current_item)
                
                print(f"\n找到的Items: {items}")
                print(f"Items数量: {len(items)}")
                
                if items:
                    self.item_combo['values'] = items
                    self.item_combo.set(items[0])
                else:
                    print("警告：未找到任何Items")
                    self.item_combo['values'] = []
                
            except Exception as e:
                print(f"更新Items时出错：{str(e)}")
                print(f"数据框列名：{self.df.columns.tolist()}")
    
    def generate_report(self):
        if not hasattr(self, 'df'):
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "请先导入历史数据！")
            return
        
        selected_area = self.area_var.get()
        selected_month = self.month_var.get()
        selected_item = self.item_var.get()
        
        if not all([selected_area, selected_month, selected_item]):
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "请选择开发领域、月份和具体Item！")
            return
        
        progress = self.progress_text.get(1.0, tk.END).strip()
        plan = self.plan_text.get(1.0, tk.END).strip()
        
        try:
            # 获取选定Item的历史数据
            area_data = self.df[self.df['Development Area'] == selected_area]
            item_number = selected_item.split(':')[0].strip()
            base_number = ''.join(filter(str.isdigit, item_number))
            
            # 使用更精确的匹配逻辑
            item_data = area_data[
                area_data['Details'].str.contains(f'tem ?{base_number}[:\n]', case=False, regex=True)
            ]
            
            # 打印调试信息
            print("\n=== 调试信息 ===")
            print(f"选择的Area: {selected_area}")
            print(f"选择的Item编号: {base_number}")
            print(f"找到的数据行数: {len(item_data)}")
            if not item_data.empty:
                row = item_data.iloc[0]
                print("\n数据行内容:")
                for col in self.df.columns:
                    print(f"{col}: {row[col]}")
            print("===============\n")
            
            # 生成提示
            prompt = ""
            if not item_data.empty:  # 只要有数据就可以生成提示
                row = item_data.iloc[0]
                
                # 月份映射表
                month_numbers = {
                    "Sep-24": 9,
                    "Oct-24": 10,
                    "Nov-24": 11,
                    "Dec-24": 12
                }
                
                # 构建历史数据字符串
                history_text = ""
                print("\n=== 历史数据构建过程 ===")
                for month in ["Sep-24", "Oct-24", "Nov-24"]:
                    # 使用数字比较月份
                    if month_numbers[month] >= month_numbers[selected_month]:
                        print(f"跳过 {month} (月份 {month_numbers[month]} 大于等于选择的月份 {month_numbers[selected_month]})")
                        continue
                        
                    progress = str(row[f"{month} Current Progress"]).strip()
                    plan = str(row[f"{month} Acting Plan"]).strip()
                    
                    print(f"\n处理 {month}:")
                    print(f"Progress: {progress}")
                    print(f"Plan: {plan}")
                    
                    # 检查是否为有效数据
                    has_progress = progress and not progress.lower() in ['nan', 'none', 'null', '']
                    has_plan = plan and not plan.lower() in ['nan', 'none', 'null', '']
                    
                    if has_progress or has_plan:
                        history_text += f"\n{month}:\n"
                        if has_progress:
                            history_text += f"Current Progress: {progress}\n"
                            print(f"添加了 {month} 的 Progress")
                        if has_plan:
                            history_text += f"Acting Plan: {plan}\n"
                            print(f"添加了 {month} 的 Plan")
                
                print("\n最终的历史数据:")
                print(history_text)
                print("===================\n")
                
                if history_text.strip():  # 确保历史数据不为空
                    if progress or plan:
                        prompt = f"""
                        Based on the following historical data from previous months:
                        {history_text}
                        
                        For {selected_month}, please generate a professional monthly report that:
                        1. Maintains the same level of detail and length as the historical records
                        2. Follows the writing style, tone and structure of previous months
                        3. Uses clear, professional language without repeating item descriptions
                        4. Shows natural progression from previous months' achievements
                        5. Uses passive voice instead of "I" or "We" as subjects
                        6. Includes specific details and examples like in historical records
                        
                        Current input to refine (if empty, generate based on historical pattern):
                        Current Progress: {progress if progress else '[Generate based on historical pattern]'}
                        Acting Plan: {plan if plan else '[Generate based on historical pattern]'}
                        
                        Note: 
                        - Match the length and detail level of historical entries
                        - Use passive voice (e.g., "The task was completed" instead of "I completed the task")
                        - Each section typically contains 3-4 sentences with specific details
                        
                        Format your response exactly as follows:
                        Current Progress:
                        [Write a detailed progress update in passive voice]
                        
                        Acting Plan:
                        [Write a detailed action plan in passive voice]
                        
                        Then provide Chinese translation for both sections:
                        
                        Current Progress (中文翻译):
                        [Translate the Current Progress to Chinese]
                        
                        Acting Plan (中文翻译):
                        [Translate the Acting Plan to Chinese]
                        """
                    else:
                        prompt = f"""
                        Based on the following historical data from previous months:
                        {history_text}
                        
                        For {selected_month}, please generate a professional monthly report that:
                        1. Maintains the same level of detail and length as the historical records
                        2. Follows the writing style, tone and structure of previous months
                        3. Uses clear, professional language without mentioning item numbers
                        4. Shows natural progression from previous months' achievements
                        5. Uses passive voice instead of "I" or "We" as subjects
                        6. Includes specific details and examples like in historical records
                        
                        Note: 
                        - Match the length and detail level of historical entries
                        - Use passive voice (e.g., "The task was completed" instead of "I completed the task")
                        - Each section typically contains 3-4 sentences with specific details
                        
                        Format your response exactly as follows:
                        Current Progress:
                        [Write a detailed progress update in passive voice]
                        
                        Acting Plan:
                        [Write a detailed action plan in passive voice]
                        
                        Then provide Chinese translation for both sections:
                        
                        Current Progress (中文翻译):
                        [Translate the Current Progress to Chinese]
                        
                        Acting Plan (中文翻译):
                        [Translate the Acting Plan to Chinese]
                        """
            
            # 只有在有提示时才调用API
            if prompt:
                client = genai.Client(api_key=self.api_key)
                response = client.models.generate_content(
                    model="gemini-2.0-flash",
                    contents=prompt
                )
                
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, response.text)
            else:
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, "无法生成提示，请检查历史数据。")
            
        except Exception as e:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"生成错误：{str(e)}")
            print(f"详细错误：{str(e)}")

def main():
    root = tk.Tk()
    app = MonthlyReportGenerator(root)
    root.mainloop()

if __name__ == '__main__':
    main() 