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
    if platform.system() == "Windows":
        return r"C:\Windows\Fonts"
    elif platform.system() == "Darwin":
        return "/Library/Fonts"
    elif platform.system() == "Linux":
        return "/usr/share/fonts"
    return ""


class ToolTip:
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
    mapping_result = {
        "text_mapping": {},
        "image_mapping": {},
        "font_mapping": {},
        "color_mapping": {},
        "font_size_mapping": {},
        "confirmed": False,
        "align_mapping": {}
    }

    dialog = tk.Toplevel(parent_window)
    dialog.title("选择映射关系")
    dialog.geometry("1050x600")
    dialog.transient(parent_window)
    dialog.grab_set()

    x = parent_window.winfo_x() + (parent_window.winfo_width() - 1050) // 2
    y = parent_window.winfo_y() + (parent_window.winfo_height() - 600) // 2
    dialog.geometry(f"+{x}+{y}")

    notebook = ttk.Notebook(dialog)
    notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    text_frame = ttk.Frame(notebook, padding=10)
    notebook.add(text_frame, text="文本、字体与颜色映射")

    ttk.Label(text_frame, text="请选择PSD文本图层与Excel列的映射关系、使用的字体和颜色:").pack(pady=5)

    text_canvas = tk.Canvas(text_frame, borderwidth=0)
    text_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_canvas.yview)
    text_scrollable_frame = ttk.Frame(text_canvas)

    text_scrollable_frame.bind(
        "<Configure>",
        lambda e: text_canvas.configure(scrollregion=text_canvas.bbox("all"))
    )

    text_canvas.create_window((0, 0), window=text_scrollable_frame, anchor="nw")
    text_canvas.configure(yscrollcommand=text_scrollbar.set)

    text_canvas.pack(side="left", fill="both", expand=True, pady=5)
    text_scrollbar.pack(side="right", fill="y")

    header_frame = ttk.Frame(text_scrollable_frame)
    header_frame.pack(fill=tk.X, pady=5)

    ttk.Label(header_frame, text="图层名称", width=10).grid(row=0, column=0, padx=5)
    ttk.Label(header_frame, text="对应Excel列", width=20).grid(row=0, column=1, padx=5)
    ttk.Label(header_frame, text="字体文件", width=43).grid(row=0, column=2, padx=5)
    ttk.Label(header_frame, text="字体大小", width=10).grid(row=0, column=3, padx=5)
    ttk.Label(header_frame, text="文本颜色", width=13).grid(row=0, column=4, padx=5)
    ttk.Label(header_frame, text="对齐方式", width=10).grid(row=0, column=5, padx=5)

    ttk.Separator(text_scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=5)

    text_comboboxes = []
    font_entries = []
    font_size_entries = []
    color_selections = {}
    color_labels = {}
    align_labels = {}

    def browse_font(entry_var):
        initial_dir = get_system_font_folder()
        font_file = filedialog.askopenfilename(
            parent=dialog,
            title="选择字体文件",
            filetypes=[("字体文件", "*.ttf *.otf *.ttc"), ("所有文件", "*.*")],
            initialdir=initial_dir
        )
        if font_file:
            entry_var.set(font_file)

    def choose_color(layer_name):
        initial_color = color_selections.get(layer_name)
        if initial_color:
            r, g, b, _ = initial_color
            hex_color = f"#{r:02x}{g:02x}{b:02x}"
        else:
            hex_color = "#000000"
        color = colorchooser.askcolor(hex_color, parent=dialog, title=f"为'{layer_name}'选择颜色")
        if color[1]:
            hex_color = color[1]
            r, g, b = tuple(int(hex_color[i:i+2], 16) for i in (1, 3, 5))
            new_color = (r, g, b, 255)
            color_selections[layer_name] = new_color
            color_labels[layer_name].config(background=hex_color)
            mapping_result["color_mapping"][layer_name] = new_color

    def set_align(layer_name, align):
        mapping_result["align_mapping"][layer_name] = align
        align_labels[layer_name].config(text=f"{align[0]} {align[1]}")

    for i, layer in enumerate(text_layers):
        row_frame = ttk.Frame(text_scrollable_frame)
        row_frame.pack(fill=tk.X, pady=3)

        layer_name = layer['name']
        layer_label = ttk.Label(row_frame, text=layer_name[:10], width=10)
        layer_label.pack(side=tk.LEFT, padx=5)
        ToolTip(layer_label, layer_name)

        data_combo = ttk.Combobox(row_frame, values=["不替换"] + excel_columns, width=20)
        data_combo.pack(side=tk.LEFT, padx=5)
        data_combo.current(0)

        font_frame = ttk.Frame(row_frame)
        font_frame.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        font_var = tk.StringVar(value="保持原始字体")
        font_entry = ttk.Entry(font_frame, textvariable=font_var, width=30)
        font_entry.pack(side=tk.LEFT, padx=2)
        font_btn = ttk.Button(font_frame, text="浏览...", width=8, command=lambda v=font_var: browse_font(v))
        font_btn.pack(side=tk.LEFT, padx=2)

        font_size_frame = ttk.Frame(row_frame)
        font_size_frame.pack(side=tk.LEFT, padx=5)
        original_size = layer.get('font_size', 12)
        if isinstance(original_size, float):
            original_size = int(original_size)
        font_size_var = tk.StringVar(value=str(original_size))
        font_size_entry = ttk.Entry(font_size_frame, textvariable=font_size_var, width=5)
        font_size_entry.pack(side=tk.LEFT)
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

        ttk.Button(size_buttons_frame, text="▲", width=2, command=lambda v=font_size_var: increase_size(v)).pack(fill=tk.X)
        ttk.Button(size_buttons_frame, text="▼", width=2, command=lambda v=font_size_var: decrease_size(v)).pack(fill=tk.X)

        color_frame = ttk.Frame(row_frame)
        color_frame.pack(side=tk.LEFT, padx=5)
        initial_color = layer.get('color', (0, 0, 0, 255))
        if initial_color:
            r, g, b = initial_color[:3]
            hex_color = f"#{r:02x}{g:02x}{b:02x}"
        else:
            hex_color = "#000000"
        color_label = tk.Label(color_frame, background=hex_color, width=3, height=1)
        color_label.pack(side=tk.LEFT)
        color_selections[layer_name] = initial_color
        color_labels[layer_name] = color_label
        color_btn = ttk.Button(color_frame, text="选择", width=8, command=lambda name=layer_name: choose_color(name))
        color_btn.pack(side=tk.LEFT, padx=2)

        align_frame = ttk.Frame(row_frame)
        align_frame.pack(side=tk.LEFT, padx=5)

        align_options = [
            ("左上", ("left", "top")),
            ("中上", ("center", "top")),
            ("右上", ("right", "top")),
            ("左中", ("left", "center")),
            ("中中", ("center", "center")),
            ("右中", ("right", "center")),
            ("左下", ("left", "bottom")),
            ("中下", ("center", "bottom")),
            ("右下", ("right", "bottom"))
        ]

        for j, (text, align) in enumerate(align_options):
            ttk.Button(align_frame, text=text, width=5,
                       command=lambda n=layer_name, a=align: set_align(n, a)).grid(row=j // 3, column=j % 3, padx=2, pady=1)

        align_label = ttk.Label(align_frame, text="左上", width=10)
        align_label.grid(row=3, column=0, columnspan=3, pady=2)
        align_labels[layer_name] = align_label
        mapping_result["align_mapping"][layer_name] = ("left", "top")

        text_comboboxes.append((layer_name, data_combo))
        font_entries.append((layer_name, font_var))
        font_size_entries.append((layer_name, font_size_var))

        info_frame = ttk.Frame(text_scrollable_frame)
        info_frame.pack(fill=tk.X, pady=0)
        font_info = f"原始字体: {layer['font']}" if layer['font'] else "原始字体: 未知"
        size_info = f"大小: {layer['font_size']}" if layer['font_size'] else "大小: 未知"
        color_info = f"颜色: {layer['color']}" if layer['color'] else "颜色: 未知"
        ttk.Label(info_frame, text=f"    {font_info}, {size_info}, {color_info}", foreground="gray").pack(side=tk.LEFT, padx=5)

        if layer['font'] and isinstance(layer['font'], dict) and 'Name' in layer['font']:
            font_name = layer['font']['Name']
            system_font_folder = get_system_font_folder()
            font_map = get_font_filename_map()
            if font_name.lower().replace(" ", "") in [k.lower().replace(" ", "") for k in font_map.keys()]:
                for k, v in font_map.items():
                    if k.lower().replace(" ", "") == font_name.lower().replace(" ", ""):
                        suggestion = os.path.join(system_font_folder, v)
                        ttk.Label(info_frame, text=f"建议字体: {suggestion}", foreground="blue").pack(side=tk.LEFT, padx=5)
                        break

        ttk.Separator(text_scrollable_frame, orient="horizontal").pack(fill=tk.X, pady=5)

    image_frame = ttk.Frame(notebook, padding=10)
    notebook.add(image_frame, text="图像映射")

    ttk.Label(image_frame, text="请选择PSD图像图层与Excel列(包含图像文件名)的映射关系:").pack(pady=5)

    image_canvas = tk.Canvas(image_frame, borderwidth=0)
    image_scrollbar = ttk.Scrollbar(image_frame, orient="vertical", command=image_canvas.yview)
    image_scrollable_frame = ttk.Frame(image_canvas)

    image_scrollable_frame.bind(
        "<Configure>",
        lambda e: image_canvas.configure(scrollregion=image_canvas.bbox("all"))
    )

    image_canvas.create_window((0, 0), window=image_scrollable_frame, anchor="nw")
    image_canvas.configure(yscrollcommand=image_scrollbar.set)

    image_canvas.pack(side="left", fill="both", expand=True, pady=5)
    image_scrollbar.pack(side="right", fill="y")

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

    ttk.Label(image_frame, text="注意: Excel中对应列应包含图像文件名(相对于数据文件夹的路径)").pack(pady=5)

    button_frame = ttk.Frame(dialog)
    button_frame.pack(pady=10)

    def confirm():
        for layer_name, combo in text_comboboxes:
            selected = combo.get()
            if selected != "不替换":
                mapping_result["text_mapping"][layer_name] = selected
        for layer_name, font_var in font_entries:
            font_path = font_var.get()
            if font_path != "保持原始字体":
                mapping_result["font_mapping"][layer_name] = font_path
        for layer_name, size_var in font_size_entries:
            try:
                font_size = int(size_var.get())
                mapping_result["font_size_mapping"][layer_name] = font_size
            except (ValueError, TypeError):
                pass
        for layer_name, combo in image_comboboxes:
            selected = combo.get()
            if selected != "不替换":
                mapping_result["image_mapping"][layer_name] = selected

        # 显示已选择的映射关系
        print("已确认的文本映射关系:")
        for layer_name, excel_col in mapping_result["text_mapping"].items():
            print(f"图层 '{layer_name}' 对应 Excel 列 '{excel_col}'")

        print("已确认的字体映射关系:")
        for layer_name, font_path in mapping_result["font_mapping"].items():
            print(f"图层 '{layer_name}' 使用字体文件 '{font_path}'")

        print("已确认的颜色映射关系:")
        for layer_name, color in mapping_result["color_mapping"].items():
            print(f"图层 '{layer_name}' 使用颜色 '{color}'")

        print("已确认的字体大小映射关系:")
        for layer_name, font_size in mapping_result["font_size_mapping"].items():
            print(f"图层 '{layer_name}' 使用字体大小 '{font_size}'")

        print("已确认的对齐方式映射关系:")
        for layer_name, align in mapping_result["align_mapping"].items():
            print(f"图层 '{layer_name}' 使用对齐方式 '{align}'")

        mapping_result["confirmed"] = True
        dialog.destroy()

    def cancel():
        mapping_result["confirmed"] = False
        dialog.destroy()

    ttk.Button(button_frame, text="确认", command=confirm).pack(side=tk.LEFT, padx=5)
    ttk.Button(button_frame, text="取消", command=cancel).pack(side=tk.LEFT, padx=5)

    parent_window.wait_window(dialog)

    if not mapping_result["confirmed"]:
        return {"text_mapping": {}, "image_mapping": {}, "font_mapping": {}, "color_mapping": {}, "font_size_mapping": {}}
    return {
        "text_mapping": mapping_result["text_mapping"],
        "image_mapping": mapping_result["image_mapping"],
        "font_mapping": mapping_result["font_mapping"],
        "color_mapping": mapping_result["color_mapping"],
        "font_size_mapping": mapping_result["font_size_mapping"],
        "align_mapping": mapping_result["align_mapping"]
    }
    

def extract_all_layers_info(psd):
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
                if hasattr(layer, 'resource_dict'):
                    if 'FontSet' in layer.resource_dict:
                        fonts = layer.resource_dict['FontSet']
                        if fonts and len(fonts) > 0:
                            text_info['font'] = fonts[0]
                if hasattr(layer, 'engine_dict'):
                    engine_data = layer.engine_dict
                    style_run = engine_data.get('StyleRun', {})
                    run_array = style_run.get('RunArray', [])
                    if run_array and len(run_array) > 0:
                        style_sheet = run_array[0].get('StyleSheet', {}).get('StyleSheetData', {})
                        if 'FontSize' in style_sheet:
                            text_info['font_size'] = style_sheet['FontSize']
                        # 获取字体颜色
                        if 'FillColor' in style_sheet:
                            color_data = style_sheet['FillColor']
                            if 'Values' in color_data:
                                values = color_data['Values']
                                print(f"Layer '{layer.name}' - Found color values: {values}")
                                # 通常RGB颜色值会存储在这里
                                if len(values) >= 3:
                                    try:
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
                text_layers.append(text_info)
            except Exception as e:
                pass
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
                pass
    return text_layers, image_layers


def render_text_with_wrapping(draw, text, rect, font, text_color, align="left", v_align="center", text_strategy="auto"):
    left, top, right, bottom = rect
    width = right - left
    height = bottom - top

    words = []
    temp_word = ""
    for char in text:
        if ord(char) > 127:
            if temp_word:
                words.append(temp_word)
                temp_word = ""
            words.append(char)
        elif char.isspace():
            if temp_word:
                words.append(temp_word)
                temp_word = ""
            words.append(" ")
        else:
            temp_word += char
    if temp_word:
        words.append(temp_word)

    original_font_size = font.size if hasattr(font, 'size') else 12
    current_font_size = original_font_size
    font_size_min = max(8, int(original_font_size * 0.6))

    final_font = font
    final_lines = []

    if text_strategy == "auto":
        while current_font_size >= font_size_min:
            if current_font_size != original_font_size:
                try:
                    if hasattr(font, 'path'):
                        final_font = ImageFont.truetype(font.path, current_font_size)
                    else:
                        final_font = font
                        break
                except:
                    final_font = font
                    break
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
            line_height = current_font_size * 1.2
            total_height = len(lines) * line_height
            if total_height <= height or current_font_size <= font_size_min:
                final_lines = lines
                break
            current_font_size -= 1
    else:
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

    line_height = final_font.size * 1.2 if hasattr(final_font, 'size') else 15
    max_lines = max(1, int(height / line_height))
    if len(final_lines) > max_lines:
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

    text_height = len(final_lines) * line_height
    if v_align == "center":
        y = top + (height - text_height) / 2
    elif v_align == "bottom":
        y = bottom - text_height
    else:
        y = top

    for line in final_lines:
        line_width = draw.textlength(line, font=final_font)
        if align == "center":
            x = left + (width - line_width) / 2
        elif align == "right":
            x = right - line_width
        else:
            x = left
        draw.text((x, y), line, font=final_font, fill=text_color)
        y += line_height


def safe_update_log(log_text, message):
    if log_text:
        log_text.after(0, lambda: log_text.configure(state="normal"))
        log_text.after(0, lambda: log_text.insert("end", message + "\n"))
        log_text.after(0, lambda: log_text.configure(state="disabled"))
        log_text.after(0, lambda: log_text.see("end"))
    print(message)


def list_available_fonts():
    system = platform.system()
    font_dirs = []
    if system == "Windows":
        font_dirs.append(r"C:\Windows\Fonts")
    elif system == "Darwin":
        font_dirs.extend(["/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts")])
    elif system == "Linux":
        font_dirs.extend(["/usr/share/fonts", "/usr/local/share/fonts", os.path.expanduser("~/.fonts")])
    font_dirs.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts"))
    found_fonts = []
    for font_dir in font_dirs:
        if os.path.exists(font_dir):
            try:
                fonts = [f for f in os.listdir(font_dir) if f.lower().endswith(('.ttf', '.ttc', '.otf'))]
                found_fonts.extend([os.path.join(font_dir, f) for f in fonts])
            except Exception as e:
                pass
    return found_fonts


def process_custom_psd(excel_file, folder_path, custom_psd_path, output_dir=None, log_text=None, parent_window=None, debug=False, text_strategy="auto"):
    try:
        safe_update_log(log_text, "正在加载PSD文件...")
        try:
            psd = PSDImage.open(custom_psd_path)
        except Exception as e:
            return f"❌ 无法打开PSD文件: {e}"

        safe_update_log(log_text, "正在提取图层信息...")
        try:
            text_layers, image_layers = extract_all_layers_info(psd)
            if len(text_layers) == 0 and len(image_layers) == 0:
                return "❌ 未在PSD中找到任何可用图层"
        except Exception as e:
            return f"❌ 提取图层信息失败: {e}"

        safe_update_log(log_text, "正在加载Excel数据...")
        try:
            df = pd.read_excel(excel_file)
            excel_columns = list(df.columns)
        except Exception as e:
            return f"❌ 读取Excel文件失败: {e}"

        safe_update_log(log_text, "请在弹出窗口中设置映射关系...")
        mapping_queue = queue.Queue()

        def show_mapping_dialog():
            mapping = create_mapping_ui(text_layers, image_layers, excel_columns, parent_window)
            mapping_queue.put(mapping)

        if parent_window:
            parent_window.after(0, show_mapping_dialog)
            mapping = mapping_queue.get()
        else:
            mapping = {"text_mapping": {}, "image_mapping": {}, "font_mapping": {}, "color_mapping": {}, "font_size_mapping": {}, "align_mapping": {}}

        text_mapping = mapping["text_mapping"]
        image_mapping = mapping["image_mapping"]
        font_mapping = mapping.get("font_mapping", {})
        color_mapping = mapping.get("color_mapping", {})
        font_size_mapping = mapping.get("font_size_mapping", {})
        align_mapping = mapping.get("align_mapping", {})

        if not text_mapping and not image_mapping:
            return "❌ 未设置任何映射关系，处理取消"

        if output_dir is None:
            output_dir = os.path.join("output", "custom_psd")
        os.makedirs(output_dir, exist_ok=True)

        debug_dir = None
        if debug:
            debug_dir = os.path.join(output_dir, "debug")
            os.makedirs(debug_dir, exist_ok=True)

        safe_update_log(log_text, f"开始处理 {len(df)} 条记录...")
        total_rows = len(df)
        for index, row in df.iterrows():
            safe_update_log(log_text, f"处理第 {index+1}/{total_rows} 条记录")
            final_image = Image.new('RGBA', psd.size, (255, 255, 255, 0))

            for layer_index, layer in enumerate(psd):
                left, top, right, bottom = layer.bbox
                layer_width = right - left
                layer_height = bottom - top

                if layer.kind == 'type' and layer.name in text_mapping:
                    try:
                        if debug:
                            safe_update_log(log_text, f"处理文本图层: '{layer.name}'")
                        excel_column = text_mapping[layer.name]
                        new_text = str(row[excel_column])
                        layer_info = next((l for l in text_layers if l['name'] == layer.name), None)

                        text_layer = Image.new("RGBA", (layer_width, layer_height), (255, 255, 255, 0))
                        draw = ImageDraw.Draw(text_layer)

                        font_size = 12
                        text_color = (0, 0, 0, 255)
                        font_name = None

                        if layer_info:
                            if layer.name in color_mapping:
                                text_color = color_mapping[layer.name]
                            elif layer_info['color']:
                                text_color = layer_info['color']
                            if layer.name in font_size_mapping:
                                font_size = font_size_mapping[layer.name]
                            elif layer_info['font_size']:
                                font_size = int(layer_info['font_size'])
                            if layer_info['font']:
                                font_name = layer_info['font']
                                if isinstance(font_name, dict) and 'Name' in font_name:
                                    font_name = font_name['Name']

                        font = None
                        font_loaded = False

                        if layer.name in font_mapping:
                            font_path = font_mapping[layer.name]
                            if os.path.exists(font_path) and font_path != "保持原始字体":
                                try:
                                    font = ImageFont.truetype(font_path, font_size)
                                    font_loaded = True
                                except Exception as e:
                                    pass
                        if not font_loaded and isinstance(font_name, str):
                            font_filename_map = get_font_filename_map()
                            font_name_lower = font_name.lower().replace(" ", "")
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
                                except Exception as e:
                                    pass
                        if not font_loaded:
                            system_fonts = []
                            if platform.system() == "Windows":
                                system_fonts = [
                                    r"C:\Windows\Fonts\msyh.ttc",
                                    r"C:\Windows\Fonts\simsun.ttc",
                                    r"C:\Windows\Fonts\simhei.ttf",
                                    r"C:\Windows\Fonts\simkai.ttf"
                                ]
                            elif platform.system() == "Darwin":
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
                            local_fonts = [
                                "fonts/SimHei.ttf", 
                                "fonts/SimSun.ttf",
                                "fonts/msyh.ttc",
                                "SimHei.ttf", 
                                "SimSun.ttf"
                            ]
                            for font_path in system_fonts + local_fonts:
                                try:
                                    font = ImageFont.truetype(font_path, font_size)
                                    font_loaded = True
                                    break
                                except:
                                    continue
                        if not font_loaded:
                            font = ImageFont.load_default()

                        if isinstance(new_text, str):
                            new_text = new_text.encode('utf-8', errors='replace').decode('utf-8')

                        align = "left"
                        v_align = "center"
                        if layer.name in align_mapping:
                            align, v_align = align_mapping[layer.name]

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

                        if debug and debug_dir:
                            debug_img = Image.new('RGBA', (layer_width + 20, layer_height + 70), (240, 240, 240, 255))
                            debug_img.paste(text_layer, (10, 10))
                            debug_draw = ImageDraw.Draw(debug_img)
                            debug_font = ImageFont.load_default()
                            debug_info = [
                                f"图层: {layer.name}",
                                f"字体: {font_mapping.get(layer.name, '默认')}",
                                f"大小: {font.size if hasattr(font, 'size') else '未知'}",
                                f"颜色: {text_color}",
                                f"对齐: {align}/{v_align}"
                            ]
                            y = layer_height + 15
                            for info in debug_info:
                                debug_draw.text((10, y), info, font=debug_font, fill=(0, 0, 0, 255))
                                y += 15
                            debug_img.save(os.path.join(debug_dir, f"debug_{index}_{layer.name}.png"), 'PNG')

                        final_image.paste(text_layer, (left, top), text_layer)

                    except Exception as e:
                        error_detail = traceback.format_exc()
                        safe_update_log(log_text, f"处理文本图层 '{layer.name}' 时出错: {str(e)}")
                        if debug:
                            safe_update_log(log_text, f"错误详情: {error_detail}")
                        layer_image = layer.topil().convert('RGBA')
                        final_image.paste(layer_image, (left, top), layer_image)

                elif hasattr(layer, 'has_pixels') and layer.has_pixels and layer.name in image_mapping:
                    try:
                        excel_column = image_mapping[layer.name]
                        image_filename = str(row[excel_column])
                        image_path = os.path.join(folder_path, image_filename.strip())
                        if os.path.exists(image_path):
                            new_image = Image.open(image_path).convert('RGBA')
                            new_image_resized = new_image.resize((layer_width, layer_height), Image.LANCZOS)
                            final_image.paste(new_image_resized, (left, top), new_image_resized)
                            if debug:
                                safe_update_log(log_text, f"处理图像图层 '{layer.name}' - 使用图片: {image_path}")
                        else:
                            safe_update_log(log_text, f"警告: 图片文件未找到: {image_path}")
                            layer_image = layer.topil().convert('RGBA')
                            final_image.paste(layer_image, (left, top), layer_image)
                    except Exception as e:
                        safe_update_log(log_text, f"处理图像图层 '{layer.name}' 时出错: {str(e)}")
                        layer_image = layer.topil().convert('RGBA')
                        final_image.paste(layer_image, (left, top), layer_image)
                else:
                    layer_image = layer.topil().convert('RGBA')
                    final_image.paste(layer_image, (left, top), layer_image)

            output_filename = os.path.join(output_dir, f"{index + 1}.png")
            final_image.save(output_filename, 'PNG')
            if index % 5 == 0 or index == total_rows - 1:
                safe_update_log(log_text, f"✅ 已完成: {index + 1}/{total_rows}")

        return f"✅ 所有图片已生成，存放在 {output_dir}"

    except Exception as e:
        error_detail = traceback.format_exc()
        print(error_detail)
        return f"❌ 处理自定义PSD时出错: {str(e)}"


def add_custom_psd_tab(notebook, parent_window):
    custom_frame = ttk.Frame(notebook, padding="10")
    notebook.add(custom_frame, text="自定义PSD")

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

    ttk.Label(custom_frame, text="文本处理策略:").grid(column=0, row=2, sticky="w", pady=5)
    text_strategy_frame = ttk.Frame(custom_frame)
    text_strategy_frame.grid(column=1, row=2, sticky="w", pady=5)
    text_strategy_var = tk.StringVar(value="auto")
    ttk.Radiobutton(text_strategy_frame, text="自动调整文字大小", variable=text_strategy_var, value="auto").pack(side=tk.LEFT, padx=(0, 10))
    ttk.Radiobutton(text_strategy_frame, text="固定文字大小(可能截断)", variable=text_strategy_var, value="fixed").pack(side=tk.LEFT)

    debug_var = tk.BooleanVar()
    debug_check = ttk.Checkbutton(custom_frame, text="启用调试模式（输出详细日志和调试图像）", variable=debug_var)
    debug_check.grid(column=1, row=3, sticky="w", pady=5)

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
        process_button.config(state="disabled")

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
                custom_frame.after(0, lambda: process_button.config(state="normal"))

        threading.Thread(target=process_thread, daemon=True).start()

    process_button = ttk.Button(custom_frame, text="开始处理", command=start_process)
    process_button.grid(column=1, row=4, pady=10)

    return custom_frame


if __name__ == "__main__":
    root = tk.Tk()
    root.title("PSD自动处理工具")
    root.geometry("800x600")

    main_notebook = ttk.Notebook(root)
    main_notebook.pack(expand=True, fill="both", padx=10, pady=10)

    add_custom_psd_tab(main_notebook, root)

    root.mainloop()