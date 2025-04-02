import tkinter as tk
from tkinter import ttk, filedialog, colorchooser
from PIL import Image, ImageDraw, ImageFont
from psd_tools import PSDImage
import os
import sys
import platform
import pandas as pd
import queue
import traceback
import threading


def get_font_filename_map():
    """返回常见字体显示名称到文件名的映射"""
    return {
        "黑体": "simhei.ttf",
        "宋体": "simsun.ttc",
        "微软雅黑": "msyh.ttc",
        "楷体": "simkai.ttf",
        "仿宋": "simfang.ttf",
        "等线": "Deng.ttf",
        "隶书": "SIMLI.TTF",
        "幼圆": "SIMYOU.TTF",
        "Arial": "arial.ttf",
        "Times New Roman": "times.ttf",
        "Courier New": "cour.ttf"
    }


def get_system_font_folder():
    """获取系统字体目录"""
    if platform.system() == "Windows":
        return r"C:\Windows\Fonts"
    elif platform.system() == "Darwin":  # macOS
        return "/Library/Fonts"
    elif platform.system() == "Linux":
        return "/usr/share/fonts"
    return ""


class ToolTip:
    """工具提示类，显示图层的完整名称"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
        
    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        # 创建工具提示窗口
        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()
        
    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


def create_mapping_ui(text_layers, image_layers, excel_columns, parent_window):
    """创建映射界面让用户选择Excel列与PSD图层的对应关系以及字体和颜色"""
    # 返回结果
    mapping_result = {
        "text_mapping": {}, 
        "image_mapping": {}, 
        "font_mapping": {},
        "color_mapping": {},
        "font_size_mapping": {},
        "confirmed": False
    }
    
    # 创建对话框窗口
    dialog = tk.Toplevel(parent_window)
    dialog.title("选择映射关系")
    dialog.geometry("850x600") 
    dialog.transient(parent_window)
    dialog.grab_set()
    
    # 确保对话框位于父窗口中央
    x = parent_window.winfo_x() + (parent_window.winfo_width() - 1050) // 2
    y = parent_window.winfo_y() + (parent_window.winfo_height() - 600) // 2
    dialog.geometry(f"+{x}+{y}")
    
    # 创建notebook
    notebook = ttk.Notebook(dialog)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # 文本映射标签页
    text_frame = ttk.Frame(notebook, padding=10)
    notebook.add(text_frame, text="文本、字体与颜色映射")
    
    ttk.Label(text_frame, text="请选择PSD文本图层与Excel列的映射关系、使用的字体和颜色:").pack(pady=5)
    
    # 创建文本映射滚动区域
    text_canvas = tk.Canvas(text_frame, borderwidth=0)
    text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_canvas.yview)
    text_scrollable_frame = ttk.Frame(text_canvas)
    
    text_scrollable_frame.bind(
        "<Configure>",
        lambda e: text_canvas.configure(
            scrollregion=text_canvas.bbox("all")
        )
    )
    
    text_canvas.create_window((0, 0), window=text_scrollable_frame, anchor="nw")
    text_canvas.configure(yscrollcommand=text_scrollbar.set)
    
    text_canvas.pack(side="left", fill="both", expand=True, pady=5)
    text_scrollbar.pack(side="right", fill="y")
    
    # 创建标题行
    header_frame = ttk.Frame(text_scrollable_frame)
    header_frame.pack(fill=tk.X, pady=5)
    
    ttk.Label(header_frame, text="图层名称", width=10).grid(row=0, column=0, padx=5)
    ttk.Label(header_frame, text="对应Excel列", width=22).grid(row=0, column=1, padx=5)
    ttk.Label(header_frame, text="字体文件", width=45).grid(row=0, column=2, padx=5)
    ttk.Label(header_frame, text="字体大小", width=10).grid(row=0, column=3, padx=5)
    ttk.Label(header_frame, text="文本颜色", width=10).grid(row=0, column=4, padx=5)
    
    # 分隔线
    ttk.Separator(text_scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=5)
    
    # 文本映射控件
    text_comboboxes = []
    font_entries = []  # 字体路径输入框
    font_size_entries = []  # 字体大小输入框
    color_selections = {}  # 保存用户选择的颜色
    color_labels = {}  # 保存颜色显示标签的引用
    
    # 字体选择函数
    def browse_font(entry_var):
        """打开字体选择对话框"""
        # 设置初始目录为系统字体目录
        initial_dir = get_system_font_folder()
        
        font_file = filedialog.askopenfilename(
            parent=dialog,
            title="选择字体文件",
            filetypes=[("字体文件", "*.ttf *.otf *.ttc"), ("所有文件", "*.*")],
            initialdir=initial_dir
        )
        
        if font_file:
            entry_var.set(font_file)
    
    # 颜色选择函数
    def choose_color(layer_name):
        """打开颜色选择器并更新颜色"""
        initial_color = color_selections.get(layer_name)
        if initial_color:
            # 将RGBA转换为十六进制格式
            r, g, b, _ = initial_color
            hex_color = f"#{r:02x}{g:02x}{b:02x}"
        else:
            hex_color = "#000000"  # 默认黑色
        
        color = colorchooser.askcolor(hex_color, parent=dialog, title=f"为'{layer_name}'选择颜色")
        if color[1]:  # 用户选择了颜色
            hex_color = color[1]
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (1, 3, 5))
            new_color = (r, g, b, 255)  # RGBA
            color_selections[layer_name] = new_color
            
            # 更新颜色显示
            color_labels[layer_name].config(background=hex_color)
            
            # 更新映射结果
            mapping_result["color_mapping"][layer_name] = new_color
    
    for i, layer in enumerate(text_layers):
        row_frame = ttk.Frame(text_scrollable_frame)
        row_frame.pack(fill=tk.X, pady=3)
        
        # 图层名称
        layer_name = layer['name']
        layer_label = ttk.Label(row_frame, text=layer_name[:10], width=10)
        layer_label.pack(side=tk.LEFT, padx=5)
        
        # 创建工具提示显示完整图层名
        ToolTip(layer_label, layer_name)
        
        # Excel列选择
        data_combo = ttk.Combobox(row_frame, values=["不替换"] + excel_columns, width=20)
        data_combo.pack(side=tk.LEFT, padx=5)
        data_combo.current(0)
        
        # 字体选择 - 使用Entry和浏览按钮
        font_frame = ttk.Frame(row_frame)
        font_frame.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        font_var = tk.StringVar(value="保持原始字体")
        font_entry = ttk.Entry(font_frame, textvariable=font_var, width=30)
        font_entry.pack(side=tk.LEFT, padx=2)
        
        font_btn = ttk.Button(
            font_frame, text="浏览...", width=8,
            command=lambda v=font_var: browse_font(v)
        )
        font_btn.pack(side=tk.LEFT, padx=2)
        
        # 字体大小设置
        font_size_frame = ttk.Frame(row_frame)
        font_size_frame.pack(side=tk.LEFT, padx=5)
        
        # 获取原始字体大小作为默认值
        original_size = layer.get('font_size', 12)
        if isinstance(original_size, float):
            original_size = int(original_size)
        
        # 创建字体大小输入
        font_size_var = tk.StringVar(value=str(original_size))
        font_size_entry = ttk.Entry(font_size_frame, textvariable=font_size_var, width=5)
        font_size_entry.pack(side=tk.LEFT)
        
        # 上下调整按钮
        size_buttons_frame = ttk.Frame(font_size_frame)
        size_buttons_frame.pack(side=tk.LEFT)
        
        def increase_size(size_var):
            try:
                current = int(size_var.get())
                size_var.set(str(current + 1))
            except ValueError:
                size_var.set("12")
                
        def decrease_size(size_var):
            try:
                current = int(size_var.get())
                if current > 1:
                    size_var.set(str(current - 1))
            except ValueError:
                size_var.set("12")
        
        ttk.Button(size_buttons_frame, text="▲", width=2, 
                  command=lambda v=font_size_var: increase_size(v)).pack(fill=tk.X)
        ttk.Button(size_buttons_frame, text="▼", width=2, 
                  command=lambda v=font_size_var: decrease_size(v)).pack(fill=tk.X)
        
        # 颜色选择按钮和颜色显示
        color_frame = ttk.Frame(row_frame)
        color_frame.pack(side=tk.LEFT, padx=5)
        
        # 初始颜色(从图层属性中获取)
        initial_color = layer.get('color', (0, 0, 0, 255))
        if initial_color:
            r, g, b = initial_color[:3]
            hex_color = f"#{r:02x}{g:02x}{b:02x}"
        else:
            hex_color = "#000000"  # 默认黑色
            initial_color = (0, 0, 0, 255)
        
        # 创建颜色显示标签
        color_label = tk.Label(color_frame, background=hex_color, width=3, height=1)
        color_label.pack(side=tk.LEFT)
        
        # 保存颜色和标签引用
        color_selections[layer_name] = initial_color
        color_labels[layer_name] = color_label
        
        # 已经初始化到映射结果中
        mapping_result["color_mapping"][layer_name] = initial_color
        
        # 颜色选择按钮
        color_btn = ttk.Button(color_frame, text="选择", width=8,
                              command=lambda name=layer_name: choose_color(name))
        color_btn.pack(side=tk.LEFT, padx=2)
        
        text_comboboxes.append((layer_name, data_combo))
        font_entries.append((layer_name, font_var))
        font_size_entries.append((layer_name, font_size_var))
        
        # 显示原始字体、大小和颜色信息
        info_frame = ttk.Frame(text_scrollable_frame)
        info_frame.pack(fill=tk.X, pady=0)
        
        font_info = f"原始字体: {layer['font']}" if layer['font'] else "原始字体: 未知"
        size_info = f"大小: {layer['font_size']}" if layer['font_size'] else "大小: 未知"
        color_info = f"颜色: {layer['color']}" if layer['color'] else "颜色: 未知"
        
        ttk.Label(info_frame, text=f"    {font_info}, {size_info}, {color_info}", 
                 foreground="gray").pack(side=tk.LEFT, padx=5)
        
        # 尝试显示字体建议
        if layer['font'] and isinstance(layer['font'], dict) and 'Name' in layer['font']:
            font_name = layer['font']['Name']
            system_font_folder = get_system_font_folder()
            font_map = get_font_filename_map()
            
            if font_name.lower().replace(" ", "") in [k.lower().replace(" ", "") for k in font_map.keys()]:
                for k, v in font_map.items():
                    if k.lower().replace(" ", "") == font_name.lower().replace(" ", ""):
                        suggestion = os.path.join(system_font_folder, v)
                        ttk.Label(info_frame, text=f"建议字体: {suggestion}", 
                                 foreground="blue").pack(side=tk.LEFT, padx=5)
                        break
        
        # 分隔线
        ttk.Separator(text_scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=5)
    
    # 图像映射标签页
    image_frame = ttk.Frame(notebook, padding=10)
    notebook.add(image_frame, text="图像映射")
    
    ttk.Label(image_frame, text="请选择PSD图像图层与Excel列(包含图像文件名)的映射关系:").pack(pady=5)
    
    # 创建图像映射滚动区域
    image_canvas = tk.Canvas(image_frame, borderwidth=0)
    image_scrollbar = ttk.Scrollbar(image_frame, orient="vertical", command=image_canvas.yview)
    image_scrollable_frame = ttk.Frame(image_canvas)
    
    image_scrollable_frame.bind(
        "<Configure>",
        lambda e: image_canvas.configure(
            scrollregion=image_canvas.bbox("all")
        )
    )
    
    image_canvas.create_window((0, 0), window=image_scrollable_frame, anchor="nw")
    image_canvas.configure(yscrollcommand=image_scrollbar.set)
    
    image_canvas.pack(side="left", fill="both", expand=True, pady=5)
    image_scrollbar.pack(side="right", fill="y")
    
    # 图像映射控件
    image_comboboxes = []
    for i, layer in enumerate(image_layers):
        row_frame = ttk.Frame(image_scrollable_frame)
        row_frame.pack(fill=tk.X, pady=3)
        
        layer_name = layer['name']
        layer_label = ttk.Label(row_frame, text=layer_name, width=15)
        layer_label.pack(side=tk.LEFT, padx=5)
        
        combo = ttk.Combobox(row_frame, values=["不替换"] + excel_columns, width=25)
        combo.pack(side=tk.RIGHT, padx=5)
        combo.current(0)
        
        image_comboboxes.append((layer_name, combo))
    
    # 添加提示信息
    ttk.Label(image_frame, text="注意: Excel中对应列应包含图像文件名(相对于数据文件夹的路径)").pack(pady=5)
    
    # 确认和取消按钮
    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)
    
    def confirm():
        # 收集文本映射
        for layer_name, combo in text_comboboxes:
            selected = combo.get()
            if selected != "不替换":
                mapping_result["text_mapping"][layer_name] = selected
        
        # 收集字体映射
        for layer_name, font_var in font_entries:
            font_path = font_var.get()
            if font_path != "保持原始字体":
                mapping_result["font_mapping"][layer_name] = font_path
        
        # 收集字体大小映射
        for layer_name, size_var in font_size_entries:
            try:
                font_size = int(size_var.get())
                mapping_result["font_size_mapping"][layer_name] = font_size
            except (ValueError, TypeError):
                # 如果用户输入的不是有效数字，则忽略
                pass
        
        # 颜色映射已经在选择颜色时更新
        
        # 收集图像映射
        for layer_name, combo in image_comboboxes:
            selected = combo.get()
            if selected != "不替换":
                mapping_result["image_mapping"][layer_name] = selected
                
        mapping_result["confirmed"] = True
        dialog.destroy()
    
    def cancel():
        mapping_result["confirmed"] = False
        dialog.destroy()
    
    ttk.Button(button_frame, text="确认", command=confirm).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT, padx=5)
    
    # 等待对话框关闭
    parent_window.wait_window(dialog)
    
    # 返回映射结果
    if not mapping_result["confirmed"]:
        return {"text_mapping": {}, "image_mapping": {}, "font_mapping": {}, "color_mapping": {}, "font_size_mapping": {}}
    return {
        "text_mapping": mapping_result["text_mapping"], 
        "image_mapping": mapping_result["image_mapping"],
        "font_mapping": mapping_result["font_mapping"],
        "color_mapping": mapping_result["color_mapping"],
        "font_size_mapping": mapping_result["font_size_mapping"]  # 返回字体大小映射
    }


def extract_all_layers_info(psd):
    """从PSD文件中提取所有图层信息，包括文本和图像图层"""
    text_layers = []
    image_layers = []
    
    for layer in psd.descendants():
        if layer.kind == 'type':
            try:
                left, top, right, bottom = layer.bbox
                text_info = {
                    'name': layer.name,
                    'text': layer.text,
                    'position': (left, top, right, bottom),
                    'font': None,
                    'font_size': None,
                    'color': None,
                    'layer': layer
                }
                
                # 尝试获取字体信息
                if hasattr(layer, 'resource_dict'):
                    # 提取字体名称
                    if 'FontSet' in layer.resource_dict:
                        fonts = layer.resource_dict['FontSet']
                        if fonts and len(fonts) > 0:
                            text_info['font'] = fonts[0]
                            print(f"Layer '{layer.name}' - Found font: {fonts[0]}")
                
                # 尝试获取文本样式信息
                if hasattr(layer, 'engine_dict'):
                    engine_data = layer.engine_dict
                    style_run = engine_data.get('StyleRun', {})
                    run_array = style_run.get('RunArray', [])
                    
                    if run_array and len(run_array) > 0:
                        style_sheet = run_array[0].get('StyleSheet', {}).get('StyleSheetData', {})
                        
                        # 获取字体大小
                        if 'FontSize' in style_sheet:
                            text_info['font_size'] = style_sheet['FontSize']
                            print(f"Layer '{layer.name}' - Found font size: {style_sheet['FontSize']}")
                        
                        # 获取字体颜色
                        if 'FillColor' in style_sheet:
                            color_data = style_sheet['FillColor']
                            if 'Values' in color_data:
                                values = color_data['Values']
                                print(f"Layer '{layer.name}' - Found color values: {values}")
                                # 通常RGB颜色值会存储在这里
                                if len(values) >= 3:
                                    try:
                                        # # 转换为RGB颜色，确保值在0-255范围内
                                        # r = max(0, min(255, int(values[0] * 255)))
                                        # g = max(0, min(255, int(values[1] * 255)))
                                        # b = max(0, min(255, int(values[2] * 255)))
                                        # text_info['color'] = (r, g, b, 255)
                                        # print(f"Layer '{layer.name}' - Converted color: {text_info['color']}")
                                        if all(0 <= v <= 1 for v in values[:3]):
                                            r = int(values[1] * 255)
                                            g = int(values[2] * 255)
                                            b = int(values[3] * 255)
                                        elif all(0 <= v <= 255 for v in values[:3]):
                                            r = int(values[1])
                                            g = int(values[2])
                                            b = int(values[3])
                                        else:
                                            print(f"Layer '{layer.name}' - Unknown color value range: {values}")
                                            r, g, b = 0, 0, 0
                                        # 确保颜色值在合法范围
                                        r = max(0, min(255, r))
                                        g = max(0, min(255, g))
                                        b = max(0, min(255, b))
                                        text_info['color'] = (r, g, b, 255)
                                        print(f"Layer '{layer.name}' - Converted color: {text_info['color']}")
                                    except Exception as e:
                                        print(f"Layer '{layer.name}' - Color conversion error: {str(e)}, values: {values}")
                
                # 打印提取的信息
                if text_info['font'] or text_info['font_size'] or text_info['color']:
                    print(f"图层 '{layer.name}' 信息提取结果:")
                    print(f"  - 字体: {text_info['font']}")
                    print(f"  - 字体大小: {text_info['font_size']}")
                    print(f"  - 颜色: {text_info['color']}")
                
                text_layers.append(text_info)
            except Exception as e:
                print(f"处理文本图层 '{layer.name}' 时出错: {str(e)}")
                traceback.print_exc()
        
        # 识别图像图层
        elif hasattr(layer, 'has_pixels') and layer.has_pixels:
            try:
                left, top, right, bottom = layer.bbox
                image_info = {
                    'name': layer.name,
                    'position': (left, top, right, bottom),
                    'size': (right - left, bottom - top),
                    'layer': layer
                }
                image_layers.append(image_info)
            except Exception as e:
                print(f"处理图像图层 '{layer.name}' 时出错: {str(e)}")
                
    return text_layers, image_layers


def render_text_with_wrapping(draw, text, rect, font, text_color, align="left", v_align="center", text_strategy="auto"):
    """使用自动换行功能渲染文本，并自适应文本大小
    
    参数:
        draw: PIL Draw对象
        text: 要渲染的文本
        rect: (left, top, right, bottom) 矩形区域
        font: PIL Font对象
        text_color: 文本颜色
        align: 水平对齐方式 'left', 'center', 'right'
        v_align: 垂直对齐方式 'top', 'center', 'bottom'
        text_strategy: 'auto'自动调整字体大小, 'fixed'固定字体大小可能截断
    """
    # 验证输入参数
    print(f"渲染文本: '{text[:30]}{'...' if len(text) > 30 else ''}'")
    print(f"  - 区域: {rect}")
    print(f"  - 字体大小: {font.size if hasattr(font, 'size') else '未知'}")
    print(f"  - 文本颜色: {text_color}")
    print(f"  - 对齐方式: {align}, 垂直对齐: {v_align}")
    print(f"  - 文本策略: {text_strategy}")
    
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top
    
    # 分词
    words = []
    temp_word = ""
    
    for char in text:
        if ord(char) > 127:  # 中文和其他非ASCII字符
            if temp_word:
                words.append(temp_word)
                temp_word = ""
            words.append(char)
        elif char.isspace():  # 空格
            if temp_word:
                words.append(temp_word)
                temp_word = ""
            words.append(" ")
        else:  # 英文和数字
            temp_word += char
    
    if temp_word:
        words.append(temp_word)
    
    # 获取原始字体大小
    original_font_size = font.size if hasattr(font, 'size') else 12
    current_font_size = original_font_size
    font_size_min = max(8, int(original_font_size * 0.6))  # 最小不低于原始大小的60%或8px
    
    # 确定最终使用的字体和行
    final_font = font
    final_lines = []
    
    # 如果是自动调整策略，尝试缩小字体使文本适合容器
    if text_strategy == "auto":
        # 尝试从原始大小开始逐渐减小
        while current_font_size >= font_size_min:
            # 如果不是原始大小，重新创建字体对象
            if current_font_size != original_font_size:
                try:
                    if hasattr(font, 'path'):
                        final_font = ImageFont.truetype(font.path, current_font_size)
                    else:
                        # 如果无法获取路径，回退到原始字体
                        final_font = font
                        break
                except:
                    # 如果创建失败，回退到原始字体
                    final_font = font
                    break
            
            # 使用当前字体计算换行
            lines = []
            line = ""
            line_width = 0
            
            for word in words:
                word_width = draw.textlength(word, font=final_font)
                
                if line_width + word_width <= width:
                    line += word
                    line_width += word_width
                else:
                    if line:
                        lines.append(line)
                    line = word
                    line_width = word_width
            
            if line:
                lines.append(line)
            
            # 计算总文本高度
            line_height = current_font_size * 1.2
            total_height = len(lines) * line_height
            
            # 如果文本高度合适或已达最小字号，使用当前结果
            if total_height <= height or current_font_size <= font_size_min:
                final_lines = lines
                break
            
            # 减小字体大小并重试
            current_font_size -= 1
    else:
        # 使用固定大小策略，不调整字体
        lines = []
        line = ""
        line_width = 0
        
        for word in words:
            word_width = draw.textlength(word, font=font)
            
            if line_width + word_width <= width:
                line += word
                line_width += word_width
            else:
                if line:
                    lines.append(line)
                line = word
                line_width = word_width
        
        if line:
            lines.append(line)
            
        final_lines = lines
        final_font = font
    
    # 如果文本行数超过容器可显示的最大行数，可能需要截断
    line_height = final_font.size * 1.2 if hasattr(final_font, 'size') else 15
    max_lines = max(1, int(height / line_height))
    
    if len(final_lines) > max_lines:
        # 如果文本行数超过容器可显示的最大行数
        if max_lines > 1:
            final_lines = final_lines[:max_lines-1]
            if len(final_lines[-1]) > 10:
                final_lines.append(final_lines[-1][:10] + "...")
            else:
                final_lines.append("...")
        else:
            if len(final_lines[0]) > 10:
                final_lines = [final_lines[0][:10] + "..."]
            else:
                final_lines = [final_lines[0]]
    
    # 计算实际文本高度
    text_height = len(final_lines) * line_height
    
    # 根据垂直对齐方式计算起始y坐标
    if v_align == "center":
        y = top + (height - text_height) / 2
    elif v_align == "bottom":
        y = bottom - text_height
    else:  # top
        y = top
    
    # 绘制文本
    for line in final_lines:
        line_width = draw.textlength(line, font=final_font)
        
        # 根据对齐方式计算x坐标
        if align == "center":
            x = left + (width - line_width) / 2
        elif align == "right":
            x = right - line_width
        else:  # left
            x = left
        
        draw.text((x, y), line, font=final_font, fill=text_color)
        y += line_height


def safe_update_log(log_text, message):
    """安全地更新日志，确保在主线程中执行"""
    if log_text:
        log_text.after(0, lambda: log_text.configure(state="normal"))
        log_text.after(0, lambda: log_text.insert("end", message + "\n"))
        log_text.after(0, lambda: log_text.configure(state="disabled"))
        log_text.after(0, lambda: log_text.see("end"))
    print(message)  # 同时在控制台输出


def list_available_fonts():
    """列出系统中可用的字体文件"""
    print("检查系统可用字体...")
    
    system = platform.system()
    font_dirs = []
    
    if system == "Windows":
        font_dirs.append(r"C:\Windows\Fonts")
    elif system == "Darwin":  # macOS
        font_dirs.extend([
            "/Library/Fonts",
            "/System/Library/Fonts",
            os.path.expanduser("~/Library/Fonts")
        ])
    elif system == "Linux":
        font_dirs.extend([
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            os.path.expanduser("~/.fonts")
        ])
    
    # 添加当前目录下的fonts文件夹
    font_dirs.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts"))
    
    found_fonts = []
    
    for font_dir in font_dirs:
        if os.path.exists(font_dir):
            print(f"在 {font_dir} 中查找字体...")
            try:
                fonts = [f for f in os.listdir(font_dir) if f.lower().endswith(('.ttf', '.ttc', '.otf'))]
                print(f"找到 {len(fonts)} 个字体文件")
                for font in fonts[:10]:  # 只显示前10个
                    print(f"  - {font}")
                if len(fonts) > 10:
                    print(f"  以及其他 {len(fonts)-10} 个字体文件")
                found_fonts.extend([os.path.join(font_dir, f) for f in fonts])
            except Exception as e:
                print(f"读取字体目录 {font_dir} 时出错: {str(e)}")
    
    return found_fonts


def process_custom_psd(excel_file, folder_path, custom_psd_path, output_dir=None, log_text=None, parent_window=None, debug=False, text_strategy="auto"):
    """处理用户自定义PSD文件"""
    try:
        # 更新日志
        safe_update_log(log_text, "正在加载PSD文件...")
        
        # 加载PSD文件
        try:
            psd = PSDImage.open(custom_psd_path)
        except Exception as e:
            return f"❌ 无法打开PSD文件: {e}"
            
        safe_update_log(log_text, "正在提取图层信息...")
        
        # 列出可用字体
        if debug:
            safe_update_log(log_text, "检查系统中可用的字体...")
            available_fonts = list_available_fonts()
            safe_update_log(log_text, f"找到 {len(available_fonts)} 个字体文件")
            
        # 提取图层信息
        try:
            text_layers, image_layers = extract_all_layers_info(psd)
            safe_update_log(log_text, f"找到 {len(text_layers)} 个文本图层和 {len(image_layers)} 个图像图层")
            
            if len(text_layers) == 0 and len(image_layers) == 0:
                return "❌ 未在PSD中找到任何可用图层"
                
            # 在调试模式下显示详细的图层信息
            if debug:
                safe_update_log(log_text, "文本图层详情:")
                for i, layer in enumerate(text_layers):
                    safe_update_log(log_text, f"  {i+1}. '{layer['name']}' - 文本: '{layer['text'][:20]}...'")
                    safe_update_log(log_text, f"     - 字体: {layer['font']}")
                    safe_update_log(log_text, f"     - 大小: {layer['font_size']}")
                    safe_update_log(log_text, f"     - 颜色: {layer['color']}")
        except Exception as e:
            return f"❌ 提取图层信息失败: {e}"
            
        safe_update_log(log_text, "正在加载Excel数据...")
        
        # 加载Excel数据
        try:
            df = pd.read_excel(excel_file)
            excel_columns = list(df.columns)
        except Exception as e:
            return f"❌ 读取Excel文件失败: {e}"
            
        safe_update_log(log_text, "请在弹出窗口中设置映射关系...")
        
        # 使用队列处理映射结果
        mapping_queue = queue.Queue()
        
        def show_mapping_dialog():
            mapping = create_mapping_ui(text_layers, image_layers, excel_columns, parent_window)
            mapping_queue.put(mapping)
        
        # 在主线程中显示映射对话框
        if parent_window:
            parent_window.after(0, show_mapping_dialog)
            # 等待映射结果
            mapping = mapping_queue.get()
        else:
            mapping = {"text_mapping": {}, "image_mapping": {}, "font_mapping": {}, "color_mapping": {}, "font_size_mapping": {}}
            
        text_mapping = mapping["text_mapping"]
        image_mapping = mapping["image_mapping"]
        font_mapping = mapping.get("font_mapping", {})
        color_mapping = mapping.get("color_mapping", {})
        font_size_mapping = mapping.get("font_size_mapping", {})  # 获取字体大小映射
        
        if debug:
            if font_mapping:
                safe_update_log(log_text, f"用户选择的字体映射: {font_mapping}")
            if color_mapping:
                safe_update_log(log_text, f"用户选择的颜色映射: {color_mapping}")
            if font_size_mapping:
                safe_update_log(log_text, f"用户选择的字体大小映射: {font_size_mapping}")
        
        if not text_mapping and not image_mapping:
            return "❌ 未设置任何映射关系，处理取消"
            
        # 确保输出目录存在
        if output_dir is None:
            output_dir = os.path.join("output", "custom_psd")
        os.makedirs(output_dir, exist_ok=True)
        
        # 如果是调试模式，创建调试目录
        debug_dir = None
        if debug:
            debug_dir = os.path.join(output_dir, "debug")
            os.makedirs(debug_dir, exist_ok=True)
        
        safe_update_log(log_text, f"开始处理 {len(df)} 条记录...")
        
        # 处理每一行Excel数据
        total_rows = len(df)
        for index, row in df.iterrows():
            safe_update_log(log_text, f"处理第 {index+1}/{total_rows} 条记录")
            
            # 创建最终图像
            final_image = Image.new('RGBA', psd.size, (255, 255, 255, 0))
            
            # 处理每个图层
            for layer_index, layer in enumerate(psd):
                # 获取图层边界
                left, top, right, bottom = layer.bbox
                layer_width = right - left
                layer_height = bottom - top
                
                # 检查是否为文本图层且在映射中
                if layer.kind == 'type' and layer.name in text_mapping:
                    try:
                        if debug:
                            safe_update_log(log_text, f"处理文本图层: '{layer.name}'")
                            
                        # 获取对应的Excel值
                        excel_column = text_mapping[layer.name]
                        new_text = str(row[excel_column])
                        
                        if debug:
                            safe_update_log(log_text, f"  - 替换文本: '{new_text[:30]}{'...' if len(new_text) > 30 else ''}'")
                        
                        # 获取该图层的字体信息
                        layer_info = next((l for l in text_layers if l['name'] == layer.name), None)
                        
                        # 创建文本图层
                        text_layer = Image.new("RGBA", (layer_width, layer_height), (255, 255, 255, 0))
                        draw = ImageDraw.Draw(text_layer)
                        
                        # 默认字体设置
                        font_size = 12
                        text_color = (0, 0, 0, 255)  # 默认黑色
                        font_name = None
                        
                        # 尝试使用原始字体信息
                        if layer_info:
                            if debug:
                                safe_update_log(log_text, f"  - 原始图层信息: 字体={layer_info['font']}, 大小={layer_info['font_size']}, 颜色={layer_info['color']}")
                            
                            # 使用用户选择的颜色或原始颜色
                            if layer.name in color_mapping:
                                text_color = color_mapping[layer.name]
                                if debug:
                                    safe_update_log(log_text, f"  - 使用用户选择的颜色: {text_color}")
                            elif layer_info['color']:
                                text_color = layer_info['color']
                                if debug:
                                    safe_update_log(log_text, f"  - 使用原始颜色: {text_color}")
                            else:
                                if debug:
                                    safe_update_log(log_text, f"  - 使用默认黑色")
                            
                            # 尝试使用用户设置的字体大小、原始字体大小或默认大小
                            if layer.name in font_size_mapping:
                                font_size = font_size_mapping[layer.name]
                                if debug:
                                    safe_update_log(log_text, f"  - 使用用户设置的字体大小: {font_size}")
                            elif layer_info['font_size']:
                                font_size = int(layer_info['font_size'])
                                if debug:
                                    safe_update_log(log_text, f"  - 使用原始字体大小: {font_size}")
                            else:
                                if debug:
                                    safe_update_log(log_text, f"  - 使用默认字体大小: {font_size}")
                            
                            # 尝试使用原始字体
                            if layer_info['font']:
                                font_name = layer_info['font']
                                if isinstance(font_name, dict) and 'Name' in font_name:
                                    font_name = font_name['Name']  # 提取字体名
                                if debug:
                                    safe_update_log(log_text, f"  - 原始字体名称: {font_name}")
                        
                        # 尝试找到合适的字体
                        font = None
                        font_loaded = False
                        
                        # 1. 首先尝试使用用户选择的字体文件路径
                        if layer.name in font_mapping:
                            font_path = font_mapping[layer.name]
                            if os.path.exists(font_path) and font_path != "保持原始字体":
                                try:
                                    font = ImageFont.truetype(font_path, font_size)
                                    font_loaded = True
                                    if debug:
                                        safe_update_log(log_text, f"  - 成功加载用户选择的字体文件: {font_path}")
                                except Exception as e:
                                    if debug:
                                        safe_update_log(log_text, f"  - 用户字体文件加载失败: {str(e)}")
                        
                        # 2. 如果用户未选择字体或加载失败，尝试通过原始字体名称找到系统字体
                        if not font_loaded and isinstance(font_name, str):
                            # 字体名称到文件的映射表
                            font_filename_map = get_font_filename_map()
                            font_name_lower = font_name.lower().replace(" ", "")
                            
                            # 尝试使用映射查找
                            mapped_font = None
                            for key, value in font_filename_map.items():
                                key_lower = key.lower().replace(" ", "")
                                if key_lower in font_name_lower or font_name_lower in key_lower:
                                    mapped_font = value
                                    break
                            
                            if mapped_font and platform.system() == "Windows":
                                try:
                                    font_path = f"C:\\Windows\\Fonts\\{mapped_font}"
                                    font = ImageFont.truetype(font_path, font_size)
                                    font_loaded = True
                                    if debug:
                                        safe_update_log(log_text, f"  - 通过原始字体名映射加载: {font_path}")
                                except Exception as e:
                                    if debug:
                                        safe_update_log(log_text, f"  - 通过原始字体名映射加载失败: {str(e)}")
                        
                        # 3. 尝试系统字体
                        if not font_loaded:
                            system_fonts = []
                            if platform.system() == "Windows":
                                system_fonts = [
                                    r"C:\Windows\Fonts\msyh.ttc",    # 微软雅黑
                                    r"C:\Windows\Fonts\simsun.ttc",  # 宋体
                                    r"C:\Windows\Fonts\simhei.ttf",  # 黑体
                                    r"C:\Windows\Fonts\simkai.ttf"   # 楷体
                                ]
                            elif platform.system() == "Darwin":  # macOS
                                system_fonts = [
                                    "/System/Library/Fonts/PingFang.ttc",
                                    "/Library/Fonts/Arial Unicode.ttf",
                                    "/System/Library/Fonts/STHeiti Light.ttc"
                                ]
                            elif platform.system() == "Linux":
                                system_fonts = [
                                    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
                                    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"
                                ]
                                
                            # 项目中的字体
                            local_fonts = [
                                "fonts/SimHei.ttf", 
                                "fonts/SimSun.ttf",
                                "fonts/msyh.ttc",
                                "SimHei.ttf", 
                                "SimSun.ttf"
                            ]
                            
                            # 尝试所有系统字体
                            for font_path in system_fonts + local_fonts:
                                try:
                                    font = ImageFont.truetype(font_path, font_size)
                                    font_loaded = True
                                    if debug:
                                        safe_update_log(log_text, f"  - 成功加载系统字体: {font_path}")
                                    break
                                except:
                                    continue
                        
                        # 4. 最后使用默认字体
                        if not font_loaded:
                            if debug:
                                safe_update_log(log_text, "  - 使用PIL默认字体")
                            font = ImageFont.load_default()
                        
                        # 确保文本是UTF-8编码
                        if isinstance(new_text, str):
                            new_text = new_text.encode('utf-8', errors='replace').decode('utf-8')
                        
                        # 确定对齐方式
                        align = "left"
                        v_align = "center"  # 默认垂直居中
                        
                        # 根据图层名称判断
                        layer_name_lower = layer.name.lower()
                        if "title" in layer_name_lower:
                            align = "center"  # 标题居中
                        elif "name" in layer_name_lower or "姓名" in layer.name:
                            v_align = "top"   # 姓名靠上对齐
                        
                        # 渲染文本
                        render_text_with_wrapping(
                            draw, 
                            new_text, 
                            (0, 0, layer_width, layer_height), 
                            font, 
                            text_color, 
                            align,
                            v_align,
                            text_strategy
                        )
                        
                        # 如果是调试模式，保存单独的文本图层图片
                        if debug and debug_dir:
                            debug_img = Image.new('RGBA', (layer_width + 20, layer_height + 70), (240, 240, 240, 255))
                            debug_img.paste(text_layer, (10, 10))
                            
                            # 添加调试信息
                            debug_draw = ImageDraw.Draw(debug_img)
                            debug_font = ImageFont.load_default()
                            debug_info = [
                                f"图层: {layer.name}",
                                f"字体: {font_mapping.get(layer.name, '默认')}",
                                f"大小: {font.size if hasattr(font, 'size') else '未知'}",
                                f"颜色: {text_color}"
                            ]
                            
                            y = layer_height + 15
                            for info in debug_info:
                                debug_draw.text((10, y), info, font=debug_font, fill=(0, 0, 0, 255))
                                y += 15
                                
                            debug_img.save(os.path.join(debug_dir, f"debug_{index}_{layer.name}.png"), 'PNG')
                        
                        # 粘贴到最终图像
                        final_image.paste(text_layer, (left, top), text_layer)
                        
                    except Exception as e:
                        error_detail = traceback.format_exc()
                        safe_update_log(log_text, f"处理文本图层 '{layer.name}' 时出错: {str(e)}")
                        if debug:
                            safe_update_log(log_text, f"错误详情: {error_detail}")
                        
                        # 使用原始图层
                        layer_image = layer.topil().convert('RGBA')
                        final_image.paste(layer_image, (left, top), layer_image)
                
                # 检查是否为图像图层且在映射中
                elif hasattr(layer, 'has_pixels') and layer.has_pixels and layer.name in image_mapping:
                    try:
                        # 获取对应的Excel值(图片文件名)
                        excel_column = image_mapping[layer.name]
                        image_filename = str(row[excel_column])
                        
                        # 构建完整的图片路径
                        image_path = os.path.join(folder_path, image_filename.strip())
                        
                        if os.path.exists(image_path):
                            # 加载新图像
                            new_image = Image.open(image_path).convert('RGBA')
                            
                            # 调整图像大小以适应图层
                            new_image_resized = new_image.resize((layer_width, layer_height), Image.LANCZOS)
                            
                            # 粘贴到最终图像
                            final_image.paste(new_image_resized, (left, top), new_image_resized)
                            
                            if debug:
                                safe_update_log(log_text, f"处理图像图层 '{layer.name}' - 使用图片: {image_path}")
                        else:
                            safe_update_log(log_text, f"警告: 图片文件未找到: {image_path}")
                            # 使用原始图层
                            layer_image = layer.topil().convert('RGBA')
                            final_image.paste(layer_image, (left, top), layer_image)
                    except Exception as e:
                        safe_update_log(log_text, f"处理图像图层 '{layer.name}' 时出错: {str(e)}")
                        # 如果出错，使用原始图层
                        layer_image = layer.topil().convert('RGBA')
                        final_image.paste(layer_image, (left, top), layer_image)
                else:
                    # 非映射图层保持原样
                    # layer_image = layer.topil().convert('RGBA')
                    # final_image.paste(layer_image, (left, top), layer_image)
                    layer_image = layer.topil()
                    if layer_image is not None:
                        layer_image = layer_image.convert('RGBA')
                        final_image.paste(layer_image, (left, top), layer_image)
                    else:
                        # 处理无法转换的情况，例如记录错误或跳过
                        safe_update_log(log_text, f"警告：无法转换图层 '{layer.name}'，将跳过该图层。")
            
            # 保存最终图像
            output_filename = os.path.join(output_dir, f"{index + 1}.png")
            final_image.save(output_filename, 'PNG')
            
            # 更新进度
            if index % 5 == 0 or index == total_rows - 1:  # 每5个文件或最后一个文件更新一次进度
                safe_update_log(log_text, f"✅ 已完成: {index + 1}/{total_rows}")
        
        return f"✅ 所有图片已生成，存放在 {output_dir}"
    
    except Exception as e:
        error_detail = traceback.format_exc()
        print(error_detail)  # 在控制台打印详细错误
        return f"❌ 处理自定义PSD时出错: {str(e)}"


def add_custom_psd_tab(notebook, parent_window):
    """添加自定义PSD功能标签页到界面"""
    custom_frame = ttk.Frame(notebook, padding="10")
    notebook.add(custom_frame, text="自定义PSD")
    
    # 创建控件
    ttk.Label(custom_frame, text="上传自定义PSD文件:").grid(column=0, row=0, sticky="w", pady=5)
    
    psd_path_var = tk.StringVar()
    psd_path_entry = ttk.Entry(custom_frame, width=40, textvariable=psd_path_var)
    psd_path_entry.grid(column=1, row=0, pady=5)
    
    def browse_psd():
        filename = filedialog.askopenfilename(filetypes=[("PSD文件", "*.psd")])
        if filename:
            psd_path_var.set(filename)
    
    ttk.Button(custom_frame, text="浏览...", command=browse_psd).grid(column=2, row=0, padx=5, pady=5)
    
    ttk.Label(custom_frame, text="数据文件夹:").grid(column=0, row=1, sticky="w", pady=5)
    
    folder_path_var = tk.StringVar()
    folder_path_entry = ttk.Entry(custom_frame, width=40, textvariable=folder_path_var)
    folder_path_entry.grid(column=1, row=1, pady=5)
    
    def browse_folder():
        folder = filedialog.askdirectory()
        if folder:
            folder_path_var.set(folder)
    
    ttk.Button(custom_frame, text="浏览...", command=browse_folder).grid(column=2, row=1, padx=5, pady=5)
    
    # 文本处理策略
    ttk.Label(custom_frame, text="文本处理策略:").grid(column=0, row=2, sticky="w", pady=5)
    
    text_strategy_frame = ttk.Frame(custom_frame)
    text_strategy_frame.grid(column=1, row=2, sticky="w", pady=5)
    
    text_strategy_var = tk.StringVar(value="auto")
    ttk.Radiobutton(text_strategy_frame, text="自动调整文字大小", variable=text_strategy_var, value="auto").pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(text_strategy_frame, text="固定文字大小(可能截断)", variable=text_strategy_var, value="fixed").pack(side=tk.LEFT)
    
    # 添加调试复选框
    debug_var = tk.BooleanVar()
    debug_check = ttk.Checkbutton(custom_frame, text="启用调试模式（输出详细日志和调试图像）", variable=debug_var)
    debug_check.grid(column=1, row=3, sticky="w", pady=5)
    
    # 日志文本框
    log_frame = ttk.LabelFrame(custom_frame, text="处理日志")
    log_frame.grid(column=0, row=5, columnspan=3, sticky="nsew", pady=10)
    custom_frame.grid_rowconfigure(5, weight=1)
    custom_frame.grid_columnconfigure(0, weight=0)
    custom_frame.grid_columnconfigure(1, weight=1)
    custom_frame.grid_columnconfigure(2, weight=0)
    
    log_text = tk.Text(log_frame, height=10, wrap="word", state="disabled")
    log_text.pack(fill="both", expand=True, side=tk.LEFT)
    
    scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill="y")
    log_text.configure(yscrollcommand=scrollbar.set)
    
    def update_log(message):
        log_text.configure(state="normal")
        log_text.insert("end", message + "\n")
        log_text.configure(state="disabled")
        log_text.see("end")
    
    def start_process():
        psd_path = psd_path_var.get()
        folder_path = folder_path_var.get()
        
        if not psd_path or not os.path.exists(psd_path):
            update_log("❌ 请选择有效的PSD文件!")
            return
        
        if not folder_path or not os.path.exists(folder_path):
            update_log("❌ 请选择有效的数据文件夹!")
            return
        
        # 查找Excel文件
        excel_file = ""
        for file in os.listdir(folder_path):
            if file.endswith(".xlsx"):
                excel_file = os.path.join(folder_path, file)
                break
        
        if not excel_file:
            update_log("❌ 数据文件夹中未找到Excel文件!")
            return
        
        update_log(f"开始处理自定义PSD: {psd_path}")
        update_log(f"使用数据: {excel_file}")
        update_log(f"文本处理策略: {text_strategy_var.get()}")
        
        # 禁用按钮，防止重复点击
        process_button.config(state="disabled")
        
        # 在新线程中运行处理函数
        def process_thread():
            try:
                result = process_custom_psd(
                    excel_file, 
                    folder_path, 
                    psd_path, 
                    log_text=log_text,
                    parent_window=parent_window,
                    debug=debug_var.get(),
                    text_strategy=text_strategy_var.get()
                )
                custom_frame.after(0, lambda: update_log(result))
            except Exception as e:
                error_detail = traceback.format_exc()
                print(error_detail)
                custom_frame.after(0, lambda: update_log(f"❌ 处理出错: {str(e)}"))
            finally:
                # 重新启用按钮
                custom_frame.after(0, lambda: process_button.config(state="normal"))
        
        threading.Thread(target=process_thread, daemon=True).start()

    # 保存按钮引用
    process_button = ttk.Button(custom_frame, text="开始处理", command=start_process)
    process_button.grid(column=1, row=4, pady=10)
    
    return custom_frame


# 如果作为独立程序运行
if __name__ == "__main__":
    # 创建主窗口
    root = tk.Tk()
    root.title("批量出图工具@乐乐大王")
    root.geometry("800x600")
    
    # 创建标签页控件
    main_notebook = ttk.Notebook(root)
    main_notebook.pack(expand=True, fill="both", padx=10, pady=10)
    
    # 添加自定义PSD标签页
    add_custom_psd_tab(main_notebook, root)
    
    # 启动主循环
    root.mainloop()