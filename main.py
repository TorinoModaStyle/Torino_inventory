import flet as ft
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import os
import datetime
import numpy as np
import shutil  # 👈 إضافة هذه المكتبة للنسخ الآمن بين الذاكرة المحلية والدرايف

# 1. تعريف المتغيرات العامة للمسارات (سيتم تعبئتها ديناميكياً داخل دالة main)
BASE_PATH = ""
MASTER_FILE = ""
LOCAL_JARD_FILE = ""

# 2. قاموس رموز الألوان الرسمي والمعتمد لشركة Torino أوفلاين
TORINO_COLORS = {
    "01": "Black", "02": "Silver", "03": "Gold", "04": "Beige", "05": "Navy",
    "06": "Gray", "07": "Red", "08": "Blue", "09": "Fuchia", "10": "White",
    "11": "Champagne", "12": "Brown", "13": "Camel", "14": "Green", "15": "Orange",
    "16": "Phosphoric", "17": "Light golden", "18": "Khaki", "19": "Maroon", "20": "Bronze",
    "21": "Pink", "22": "Purple", "23": "Light Beige", "24": "OLIVE", "25": "Yellow",
    "26": "Rainbow", "27": "Platinum", "28": "Dark Brown", "29": "Copper", "30": "Bronette",
    "31": "Espresso", "32": "Red Brown", "33": "Off White", "34": "Coffee", "35": "Dark Green",
    "36": "Nude", "37": "Tan", "38": "Pewter", "39": "Light Green", "40": "MUD",
    "41": "Light Gray", "42": "Apricot", "43": "Dark Pink", "44": "Light Black", "45": "Army Green",
    "46": "Cream", "47": "Peach", "48": "F.Green", "49": "CORAL", "50": "Dark Navy",
    "51": "Dark Grey", "52": "L.Blue", "53": "Rose Gold", "54": "Wine", "55": "Dark Gold",
    "56": "Sky Blue", "57": "Light Purple", "58": "Antique Gold", "59": "Antique Silver", "60": "Cuticolor",
    "61": "D-Black", "62": "D-Coffee", "63": "D-Maroon", "64": "D-Purple", "65": "L-Brown",
    "66": "L-Pink", "67": "L-Yellow", "68": "Silvery Grey", "69": "Turquoise", "70": "Yellow Brown",
    "71": "Stone", "72": "S.BLACK", "73": "D Red", "74": "Black/White", "75": "Almond",
    "76": "K-serpentine", "77": "D Beige", "78": "D Pink", "79": "TIFFANY", "80": "L KHAKI",
    "81": "D KHAKI", "82": "GUN", "83": "multicolour", "84": "WHITE/BLACK", "85": "BLACK/CAMEL",
    "86": "CAMEL/BLACK", "87": "WHITE/CAMEL", "88": "D- BLUE", "89": "GRNNY", "90": "LEOPARD",
    "91": "L.GREY", "92": "L.COFFEE", "93": "COLOURFULL", "94": "TIGER", "95": "ZEBRA"
}

# 3. المتغيرات العامة للتحكم بحالة النظام
current_employee = ""
master_df = pd.DataFrame()
current_jard_df = pd.DataFrame()

def clean_value(val):
    """تنظيف القيم المرجعية وحظر كتابة nan أو المسميات الفارغة تماماً للبحث النظيف"""
    if pd.isna(val) or str(val).strip().lower() in ['nan', 'none', '-', '']:
        return ""
    # تحويل الأرقام التي تحتوي على فواصل عشرية إلى رقم صحيح ثابت
    if isinstance(val, (int, float)):
        return str(int(val))
    try:
        if float(val).is_integer():
            return str(int(float(val)))
    except:
        pass
    return str(val).strip()

def parse_scanned_barcode(raw_barcode):
    """تفكيك الباركود الممسوح ميدانياً إلى مكوناته الثلاثة بناءً على عدد خاناته"""
    raw_str = str(raw_barcode).strip()
    length = len(raw_str)
    
    base_code = raw_str
    color_name = ""
    size_val = ""
    
    if length >= 10:
        base_code = raw_str[:6]
        color_code = raw_str[6:8]
        size_code = raw_str[8:10]
        color_name = TORINO_COLORS.get(color_code, "")
        size_val = clean_value(size_code)
    elif length == 8:
        base_code = raw_str[:6]
        color_code = raw_str[6:8]
        color_name = TORINO_COLORS.get(color_code, "")
    elif length == 6:
        base_code = raw_str
        
    return base_code, color_name, size_val
def init_offline_database():
    """تهيئة وقراءة ملف الماستر داتا وملف الجرد التراكمي محلياً أوفلاين وبشكل نظيف"""
    global master_df, current_jard_df
    
    # أ) قراءة وتنظيف ملف الماستر داتا المرجعي الأساسي
    if os.path.exists(MASTER_FILE):
        try:
            master_df = pd.read_excel(MASTER_FILE)
            master_df.columns = [str(c).strip() for c in master_df.columns]
            # تنظيف وتوحيد نصوص الأعمدة لمنع تشتت الفلاتر
            for col in master_df.columns:
                master_df[col] = master_df[col].apply(clean_value)
        except:
            master_df = pd.DataFrame(columns=['رمز المادة', 'الباركود', 'اللون', 'القياس', 'الكمية_الأصلية'])
    else:
        master_df = pd.DataFrame(columns=['رمز المادة', 'الباركود', 'اللون', 'القياس', 'الكمية_الأصلية'])

    # ب) قراءة وتنظيف سجل الجرد التراكمي الحالي المخزن على الهاتف
    if os.path.exists(LOCAL_JARD_FILE):
        try:
            current_jard_df = pd.read_excel(LOCAL_JARD_FILE)
            for col in current_jard_df.columns:
                current_jard_df[col] = current_jard_df[col].apply(clean_value)
        except:
            current_jard_df = pd.DataFrame(columns=['رمز المادة', 'الرمز / الباركود', 'اللون', 'القياس', 'الأصلية', 'المجرودة', 'الفرق'])
    else:
        current_jard_df = pd.DataFrame(columns=['رمز المادة', 'الرمز / الباركود', 'اللون', 'القياس', 'الأصلية', 'المجرودة', 'الفرق'])
def get_row_colors(diff):
    """حساب الألوان الذكية اللحظية بدقة بناءً على منطق المطابقة لشركة Torino"""
    if diff < 0:
        # حالة العجز: خلفية حمراء فاتحة ونصوص داكنة واضحة
        return "#FFCDD2", "#B71C1C"
    elif diff == 0:
        # حالة المطابقة التامة: خلفية صفراء فاتحة ونصوص صفراء داكنة للمطابقة
        return "#FFF9C4", "#F57F17"
    else:
        # حالة الزيادة أو مادة جديدة: خلفية خضراء فاتحة ونصوص خضراء داكنة للفائض
        return "#C8E6C9", "#1B5E20"

def create_table_cell(text_val, bg_col, text_col, is_bold=False):
    """بناء حاوية خلية جدول مرنة ونظيفة تمنع تداخل الكلمات على شاشات الهاتف"""
    return ft.DataCell(
        ft.Container(
            content=ft.Text(
                str(text_val), 
                color=text_col, 
                weight=ft.FontWeight.BOLD if is_bold else ft.FontWeight.NORMAL,
                size=12,
                text_align=ft.TextAlign.CENTER
            ),
            bgcolor=bg_col,
            padding=ft.padding.all(6),
            alignment="center",
            border_radius=4
        )
    )
# [إضافة هذا المتغير في بداية الملف أو فوق الدالة مباشرة]
last_scanned_barcode = ""

def rebuild_jard_table_rows():
    """تحديث وتلوين وفلترة سطور الجدول لحظياً بالبحث التفكيكي الذكي (رمز، باركود، لون، قياس)"""
    global last_scanned_barcode
    if jard_table is None:
        return
        
    jard_table.rows.clear()
    
    try:
        current_view = view_mode_radio.value if view_mode_radio.value else "all"
    except:
        current_view = "all"
        
    search_query = str(search_field.value).strip().upper() if search_field and search_field.value else ""
    
    df_to_show = current_jard_df.copy()
    
    # 1. تطبيق محرك البحث التفكيكي والذكي لشركة Torino
    if search_query:
        clean_query = search_query.replace("#", "")
        
        # تفكيك نص البحث لمعرفة ما إذا كان يحمل لون أو قياس خفي
        s_base, s_color, s_size = parse_scanned_barcode(clean_query)
        
        if s_color and s_size:
            # فلترة كاملة وصارمة مع اللون والقياس معاً
            df_to_show = df_to_show[
                (df_to_show['الرمز / الباركود'].astype(str).str.upper() == s_base) & 
                (df_to_show['اللون'].astype(str).str.upper() == s_color.upper()) & 
                (df_to_show['القياس'].astype(str).str.upper() == s_size.upper())
            ]
        elif s_color:
            # فلترة مرنة مع اللون فقط
            df_to_show = df_to_show[
                (df_to_show['الرمز / الباركود'].astype(str).str.upper() == s_base) & 
                (df_to_show['اللون'].astype(str).str.upper() == s_color.upper())
            ]
        else:
            # فلترة عامة وشاملة للرموز المتشابهة أو الكود النقي بدون ألوان وقياسات
            df_to_show = df_to_show[
                df_to_show['الرمز / الباركود'].astype(str).str.upper().str.contains(clean_query) |
                df_to_show['رمز المادة'].astype(str).str.upper().str.contains(clean_query) |
                df_to_show['اللون'].astype(str).str.upper().str.contains(clean_query)
            ]
    elif current_view == "last":
        # عزل وإظهار السطر الخاص بآخر باركود ممسوح حياً حالياً
        if last_scanned_barcode:
            sc_base, sc_color, sc_size = parse_scanned_barcode(last_scanned_barcode)
            df_to_show = df_to_show[
                (df_to_show['الرمز / الباركود'] == sc_base) & 
                (df_to_show['اللون'] == sc_color) & 
                (df_to_show['القياس'] == sc_size)
            ]
        else:
            df_to_show = pd.DataFrame(columns=current_jard_df.columns)
    else:
        # صمام أمان السرعة الفائقة لـ Torino لمنع تجميد الرسوميات عند البداية
        try:
            df_to_show = df_to_show[df_to_show['المجرودة'].astype(float) > 0]
        except:
            df_to_show = pd.DataFrame(columns=current_jard_df.columns)
        
    # 2. رسم وبناء السطور المفلترة والمطلوبة فقط بسرعة البرق دون أي فراغات بيضاء
    for idx, row in df_to_show.iterrows():
        orig = int(float(row.get('الأصلية', 0))) if row.get('الأصلية', '') else 0
        jard = int(float(row.get('المجرودة', 0))) if row.get('المجرودة', '') else 0
        diff = int(float(row.get('الفرق', 0))) if row.get('الفرق', '') else 0
        
        bg_color, text_color = get_row_colors(diff)
        
        jard_table.rows.append(
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(str(row.get('رمز المادة', '')), color=text_color, weight=ft.FontWeight.W_500, bgcolor=bg_color)),
                    ft.DataCell(ft.Text(str(row.get('الرمز / الباركود', '')).upper(), color=text_color, weight=ft.FontWeight.BOLD, bgcolor=bg_color)),
                    ft.DataCell(ft.Text(str(row.get('اللون', '')), color=text_color, bgcolor=bg_color)),
                    ft.DataCell(ft.Text(str(row.get('القياس', '')), color=text_color, bgcolor=bg_color)),
                    ft.DataCell(ft.Text(str(orig) if orig != 0 or str(row.get('اللون', '')) != 'مستجد / جديد' else "0", color=text_color, bgcolor=bg_color)),
                    ft.DataCell(ft.Text(str(jard), color=text_color, bgcolor=bg_color)),
                    ft.DataCell(ft.Text(str(diff), color=text_color, weight=ft.FontWeight.BOLD, bgcolor=bg_color))
                ]
            )
        )
    try:
        page.update()
    except:
        pass

def parse_scanned_barcode(raw_barcode):
    """تفكيك الباركود الممسوح ميدانياً إلى مكوناته الثلاثة مع تحويله تلقائياً لأحرف كبيرة"""
    raw_str = str(raw_barcode).strip().upper() # إجبار الحروف على أن تكون كبيرة فوراً
    length = len(raw_str)
    
    base_code = raw_str
    color_name = ""
    size_val = ""
    
    if length >= 10:
        base_code = raw_str[:6]
        color_code = raw_str[6:8]
        size_code = raw_str[8:10]
        color_name = TORINO_COLORS.get(color_code, "")
        size_val = clean_value(size_code)
    elif length == 8:
        base_code = raw_str[:6]
        color_code = raw_str[6:8]
        color_name = TORINO_COLORS.get(color_code, "")
    elif length == 6:
        base_code = raw_str
        
    return base_code, color_name, size_val

def process_barcode_scan(raw_barcode):
    """المحرك الأسرع والأكثر مرونة للمطابقة الثلاثية الفورية داخل ذاكرة الجهاز أوفلاين"""
    global current_jard_df
    if not raw_barcode:
        return
        
    scanned_base, scanned_color, scanned_size = parse_scanned_barcode(raw_barcode)
    
    # تنظيف الفراغات الخفية تماماً من كافة السجلات المحلية قبل إجراء المقارنة بالذاكرة
    if not current_jard_df.empty:
        current_jard_df['الرمز / الباركود'] = current_jard_df['الرمز / الباركود'].astype(str).str.strip().str.upper()
        current_jard_df['اللون'] = current_jard_df['اللون'].astype(str).str.strip()
        current_jard_df['القياس'] = current_jard_df['القياس'].astype(str).str.strip()

    # محرك المطابقة الثلاثية الذكي والسريع
    found_idx = current_jard_df[
        (current_jard_df['الرمز / الباركود'] == scanned_base) & 
        (current_jard_df['اللون'] == scanned_color) & 
        (current_jard_df['القياس'] == scanned_size)
    ].index
    
    if len(found_idx) > 0:
        idx = found_idx[0]
        orig_val = int(float(current_jard_df.at[idx, 'الأصلية'])) if current_jard_df.at[idx, 'الأصلية'] else 0
        try:
            current_jard = int(float(current_jard_df.at[idx, 'المجرودة']))
        except:
            current_jard = 0
        new_jard = current_jard + 1
        current_jard_df.at[idx, 'المجرودة'] = str(new_jard)
        current_jard_df.at[idx, 'الفرق'] = str(new_jard - orig_val)
    else:
        # البحث المباشر بداخل ملف الماستر داتا مع تطهير الفراغات
        if not master_df.empty:
            master_df['الباركود'] = master_df['الباركود'].astype(str).str.strip().str.upper()
            master_df['اللون'] = master_df['اللون'].astype(str).str.strip()
            master_df['القياس'] = master_df['القياس'].astype(str).str.strip()
            
        master_match = master_df[
            (master_df['الباركود'] == scanned_base) & 
            (master_df['اللون'] == scanned_color) & 
            (master_df['القياس'] == scanned_size)
        ]
        
        if not master_match.empty:
            m_row = master_match.iloc[0]
            orig_qty = int(float(m_row.get('الكمية_الأصلية', 0))) if m_row.get('الكمية_الأصلية', '') else 0
            new_row = {
                'رمز المادة': str(m_row.get('رمز المادة', '')),
                'الرمز / الباركود': scanned_base, 'اللون': scanned_color, 'القياس': scanned_size,
                'الأصلية': str(orig_qty), 'المجرودة': "1", 'الفرق': str(1 - orig_qty)
            }
            current_jard_df = pd.concat([current_jard_df, pd.DataFrame([new_row])], ignore_index=True)
        else:
            # إضافة سطر جديد للمستجدات الفائضة بالأسفل
            new_row = {
                'رمز المادة': "", 'الرمز / الباركود': scanned_base,
                'اللون': scanned_color if scanned_color else "مستجد / جديد",
                'القياس': scanned_size if scanned_size else "-",
                'الأصلية': "0", 'المجرودة': "1", 'الفرق': "1"
            }
            current_jard_df = pd.concat([current_jard_df, pd.DataFrame([new_row])], ignore_index=True)
            
    try:
        current_jard_df.to_excel(LOCAL_JARD_FILE, index=False)
    except:
        pass

    # [ضع هذا السطر قبل سحب التحديث بالأسفل مباشرة داخل دالة process_barcode_scan]
    global last_scanned_barcode
    last_scanned_barcode = str(raw_barcode).strip().upper()
    
    save_individual_scan_with_timestamp(raw_barcode)
    rebuild_jard_table_rows()

def save_individual_scan_with_timestamp(barcode_str):
    """حفظ طابع الوقت والثانية للمسحة الحالية على القرص فوراً دون إحداث بطء في النظام الميداني"""
    global current_employee
    emp_name = current_employee if current_employee else "عام"
    emp_file = os.path.join(BASE_PATH, f"scans_{emp_name}.xlsx")
    
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    new_scan = {
        'الباركود المقروء': str(barcode_str).strip().upper(),
        'الكمية الممسوحة': 1,
        'تاريخ ووقت القراءة المباشر': now_str
    }
    
    # حفظ مباشر وسريع لملف الطابع الزمني لضمان ثباته التاريخي
    try:
        df = pd.read_excel(emp_file) if os.path.exists(emp_file) else pd.DataFrame()
        df = pd.concat([df, pd.DataFrame([new_scan])], ignore_index=True)
        df.to_excel(emp_file, index=False)
    except:
        pass

def apply_quantity_correction(e):
    """تعديل كمية المادة بسرعة البرق مع فلترة وعزل السطر المطلوب فوراً لمنع التشتت والبطء"""
    global current_jard_df
    barcode_str = str(edit_barcode_field.value).strip().upper()
    
    try:
        new_qty = int(edit_qty_field.value)
    except:
        new_qty = 0
        
    if not barcode_str:
        return
        
    # 1. تفكيك الباركود بدقة للمطابقة الثلاثية الصارمة لشركة Torino
    sc_base, sc_color, sc_size = parse_scanned_barcode(barcode_str)
    
    if not current_jard_df.empty:
        current_jard_df['الرمز / الباركود'] = current_jard_df['الرمز / الباركود'].astype(str).str.strip().str.upper()
        current_jard_df['اللون'] = current_jard_df['اللون'].astype(str).str.strip()
        current_jard_df['القياس'] = current_jard_df['القياس'].astype(str).str.strip()

    # 2. البحث السريع جداً في الذاكرة العشوائية (RAM)
    found_idx = current_jard_df[
        (current_jard_df['الرمز / الباركود'] == sc_base) & 
        (current_jard_df['اللون'] == sc_color) & 
        (current_jard_df['القياس'] == sc_size)
    ].index
    
    if len(found_idx) > 0:
        idx = found_idx[0]
        orig_val = int(float(current_jard_df.at[idx, 'الأصلية'])) if current_jard_df.at[idx, 'الأصلية'] else 0
        current_jard_df.at[idx, 'المجرودة'] = str(new_qty)
        current_jard_df.at[idx, 'الفرق'] = str(new_qty - orig_val)
        
        # 3. [الفلترة الفورية] حقن الباركود في خانة البحث لعزل السطر المطلوب فوراً أمام عين الموظف
        search_field.value = barcode_str
        
        # 4. تحديث الجدول صامتاً بناءً على الفلترة الجديدة الفائقة السرعة
        rebuild_jard_table_rows()
        
        # حفظ الكاش المحلي أوفلاين صامتاً في الخلفية دون تعطيل الواجهة
        try:
            current_jard_df.to_excel(LOCAL_JARD_FILE, index=False)
        except:
            pass
            
        edit_barcode_field.value = ""
        edit_qty_field.value = "0"
        
        status_text.value = f"✅ تم عزل المادة وتعديل كمية {barcode_str} إلى {new_qty} بسرعة فائقة!"
        status_text.color = "#008000"
    else:
        status_text.value = "❌ عذراً، الباركود المدخل للتعديل غير موجود في مسحات الجرد الحالي!"
        status_text.color = "#850000"
        
    page.update()

def trigger_export_excel(e):
    """المحرك الأكبر لبناء ملف الإكسل الختامي الملون المجمع وحفظه محلياً ثم مشاركته"""
    global current_jard_df, BASE_PATH
    
    # 1. توليد اسم فريد ونظيف بالوقت والثانية لحفظ الملف محلياً في المجلد المضمون للأندرويد
    timestamp_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    FINAL_EXPORT_FILE = os.path.join(BASE_PATH, f"Torino_Report_{timestamp_str}.xlsx")

    try:
        # إنشاء ملف إكسل فارغ والبدء في تعبئة الصفحات باستخدام openpyxl
        wb = openpyxl.Workbook()
        
        # --- الصفحة الأولى: الجرد الكامل الموحد ---
        ws1 = wb.active
        ws1.title = "الجرد الكامل الموحد"
        ws1.sheet_view.showGridLines = True
        
        # كتابة العناوين الأساسية الشاملة
        headers1 = ['رمز المادة', 'الباركود', 'اللون', 'القياس', 'الكمية_الأصلية', 'الكمية_المجرودة', 'الفرق']
        ws1.append(headers1)
        
        # تعبئة البيانات وصبغ السطور بالألوان الذكية
        for _, row in current_jard_df.iterrows():
            orig = int(float(row.get('الأصلية', 0))) if row.get('الأصلية', '') else 0
            jard = int(float(row.get('المجرودة', 0))) if row.get('المجرودة', '') else 0
            diff = int(float(row.get('الفرق', 0))) if row.get('الفرق', '') else 0
            
            if diff < 0:
                fill_color = "FFCDD2"  # أحمر العجز
            elif diff == 0:
                fill_color = "FFF9C4"  # أصفر المطابقة
            else:
                fill_color = "C8E6C9"  # أخضر الزيادة
                
            row_data = [
                row.get('رمز المادة', ''), row.get('الرمز / الباركود', ''),
                row.get('اللون', ''), row.get('القياس', ''), orig, jard, diff
            ]
            ws1.append(row_data)
            
            # صبغ وتوسيط خلايا السطر الحالي بالكامل
            for cell in ws1[ws1.max_row]:
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # --- الصفحة الثانية: الفروقات والأخطاء ---
        ws2 = wb.create_sheet(title="الفروقات والأخطاء")
        ws2.sheet_view.showGridLines = True
        
        ws2.cell(row=1, column=1, value="جدول النقص (الجانب الأيمن)").font = Font(bold=True, size=12, color="B71C1C")
        ws2.cell(row=1, column=9, value="جدول الزيادة والمستجدات (الجانب الأيسر)").font = Font(bold=True, size=12, color="1B5E20")
        
        sub_headers = ['رمز المادة', 'الباركود', 'اللون', 'القياس', 'الكمية_الأصلية', 'الكمية_المجرودة', 'الفرق']
        for i, h in enumerate(sub_headers):
            ws2.cell(row=2, column=i+1, value=h)
            ws2.cell(row=2, column=i+9, value=h)
            
        df_neg = current_jard_df[current_jard_df['الفرق'].astype(int) < 0]
        df_pos = current_jard_df[current_jard_df['الفرق'].astype(int) > 0]
        
        r_idx = 3
        for _, row in df_neg.iterrows():
            ws2.cell(row=r_idx, column=1, value=row.get('رمز المادة', ''))
            ws2.cell(row=r_idx, column=2, value=row.get('الرمز / الباركود', ''))
            ws2.cell(row=r_idx, column=3, value=row.get('اللون', ''))
            ws2.cell(row=r_idx, column=4, value=row.get('القياس', ''))
            ws2.cell(row=r_idx, column=5, value=int(float(row.get('الأصلية', 0))))
            ws2.cell(row=r_idx, column=6, value=int(float(row.get('المجرودة', 0))))
            ws2.cell(row=r_idx, column=7, value=int(float(row.get('الفرق', 0))))
            for c in range(1, 8):
                ws2.cell(row=r_idx, column=c).fill = PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid")
                ws2.cell(row=r_idx, column=c).alignment = Alignment(horizontal="center")
            r_idx += 1
            
        r_idx = 3
        for _, row in df_pos.iterrows():
            ws2.cell(row=r_idx, column=9, value=row.get('رمز المادة', ''))
            ws2.cell(row=r_idx, column=10, value=row.get('الرمز / الباركود', ''))
            ws2.cell(row=r_idx, column=11, value=row.get('اللون', ''))
            ws2.cell(row=r_idx, column=12, value=row.get('القياس', ''))
            ws2.cell(row=r_idx, column=13, value=int(float(row.get('الأصلية', 0))))
            ws2.cell(row=r_idx, column=14, value=int(float(row.get('المجرودة', 0))))
            ws2.cell(row=r_idx, column=15, value=int(float(row.get('الفرق', 0))))
            for c in range(9, 16):
                ws2.cell(row=r_idx, column=c).fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
                ws2.cell(row=r_idx, column=c).alignment = Alignment(horizontal="center")
            r_idx += 1

        # --- الصفحة 3 و 4 و 5... تتبع سجلات كل موظف ---
        all_files = os.listdir(BASE_PATH)
        for filename in all_files:
            if filename.startswith("scans_") and filename.endswith(".xlsx"):
                emp_name = filename.replace("scans_", "").replace(".xlsx", "")
                try:
                    emp_df = pd.read_excel(os.path.join(BASE_PATH, filename))
                    ws_emp = wb.create_sheet(title=f"جرد الموظف - {emp_name}")
                    ws_emp.sheet_view.showGridLines = True
                    ws_emp.cell(row=1, column=1, value=f"📋 سجل المسحات التفصيلي للموظف: {emp_name}").font = Font(bold=True, size=12)
                    ws_emp.append(['الباركود المقروء', 'الكمية الممسوحة', 'تاريخ ووقت القراءة المباشر'])
                    
                    for _, s_row in emp_df.iterrows():
                        ws_emp.append([str(s_row.get('الباركود المقروء', '')), 1, str(s_row.get('تاريخ ووقت القراءة المباشر', ''))])
                        for c in range(1, 4):
                            ws_emp[ws_emp.max_row][c-1].alignment = Alignment(horizontal="center")
                except:
                    pass

        # حفظ ملف التقرير الختامي في الكاش المحلي بأمان
        wb.save(FINAL_EXPORT_FILE)
        
        # استدعاء نافذة المشاركة الرسمية للأندرويد لرفع الملف يدوياً إلى قوقل درايف أو إرساله
        page.share_file(FINAL_EXPORT_FILE)
        
        status_text.value = "📥 تم توليد التقرير بنجاح، يرجى اختيار Google Drive من قائمة المشاركة لحفظه!"
        status_text.color = "#008000"
    except Exception as ex:
        status_text.value = f"❌ فشل تصدير التقرير النهائي: {ex}"
        status_text.color = "#850000"
        
    page.update()

def trigger_manual_sync(e):
    """تفعيل اختيار ملف الموظف الآخر لدمج وتجميع الكميات المجرودة تراكمياً"""
    status_text.value = "🔍 يرجى اختيار ملف إكسل الخاص بالموظف الآخر لدمجه..."
    status_text.color = "#1B2A47"
    page.update()
    
    # استدعاء ملتقط الملفات المدمج لفتح واجهة نظام الأندرويد واختيار ملف إكسل محدد
    sync_file_picker.pick_files(
        allow_multiple=False,
        allowed_extensions=["xlsx"]
    )

def handle_sync_file_result(e: ft.FilePickerResultEvent):
    """المحرك الذكي لقراءة ملف الموظف المختار ودمج قراءاته في الذاكرة الحية فوراً"""
    global current_jard_df, LOCAL_JARD_FILE
    
    # صمام أمان ثنائي ومرن للتحقق من وجود الملف المستلم للأندرويد
    if not e.files or len(e.files) == 0 or e.files[0].path is None:
        status_text.value = "💡 تم إلغاء عملية المزامنة ولم يتم اختيار أي ملف."
        status_text.color = "#850000"
        page.update()
        return
        
    selected_file_path = e.files[0].path
    status_text.value = "⚡ جاري قراءة وتجميع قراءات الملف المحدد..."
    status_text.color = "#1B2A47"
    page.update()
    
    try:
        # 1. قراءة وتنظيف ملف الجرد القادم من الجهاز الآخر
        other_df = pd.read_excel(selected_file_path)
        
        if other_df.empty:
            status_text.value = "⚠️ الملف المختار فارغ ولا يحتوي على بيانات جرد!"
            status_text.color = "#850000"
            page.update()
            return
            
        # 2. محرك التجميع الثلاثي الذكي داخل الذاكرة العشوائية (RAM)
        merge_count = 0
        for _, row in other_df.iterrows():
            b_code = str(row.get('الرمز / الباركود', '')).strip().upper()
            c_code = str(row.get('اللون', '')).strip()
            s_code = str(row.get('القياس', '')).strip()
            
            try:
                j_qty = int(float(row.get('الكمية_المجرودة', row.get('المجرودة', 0))))
            except:
                j_qty = 0
                
            if j_qty > 0:
                # البحث عن تطابق السجل في الكاش الحالي للجهاز
                match_idx = current_jard_df[
                    (current_jard_df['الرمز / الباركود'].astype(str).str.upper() == b_code) &
                    (current_jard_df['اللون'].astype(str) == c_code) &
                    (current_jard_df['القياس'].astype(str) == s_code)
                ].index
                
                if len(match_idx) > 0:
                    c_idx = match_idx[0]
                    try:
                        old_qty = int(float(current_jard_df.at[c_idx, 'المجرودة']))
                    except:
                        old_qty = 0
                    
                    # دمج الكميات تراكمياً وإعادة حساب الفرق
                    new_qty = old_qty + j_qty
                    current_jard_df.at[c_idx, 'المجرودة'] = str(new_qty)
                    
                    try:
                        orig_val = int(float(current_jard_df.at[c_idx, 'الأصلية']))
                    except:
                        orig_val = 0
                    current_jard_df.at[c_idx, 'الفرق'] = str(new_qty - orig_val)
                    merge_count += 1
                else:
                    # إذا كانت مادة مستجدة بالكامل وغير موجودة في كاش هذا الجهاز
                    new_row = {
                        'رمز المادة': str(row.get('رمز المادة', '')),
                        'الرمز / الباركود': b_code, 'اللون': c_code, 'القياس': s_code,
                        'الأصلية': str(row.get('الكمية_الأصلية', row.get('الأصلية', '0'))),
                        'المجرودة': str(j_qty),
                        'الفرق': str(row.get('الفرق', j_qty))
                    }
                    current_jard_df = pd.concat([current_jard_df, pd.DataFrame([new_row])], ignore_index=True)
                    merge_count += 1

        # 3. حفظ الكاش المدمج الجديد على الهاتف وتحديث شاشة العرض بلحظتها
        if merge_count > 0:
            current_jard_df.to_excel(LOCAL_JARD_FILE, index=False)
            rebuild_jard_table_rows()
            status_text.value = f"✅ تم دمج وتجميع {merge_count} مادة بنجاح وتحديث السجلات!"
            status_text.color = "#008000"
        else:
            status_text.value = "💡 لم يتم العثور على أي كميات مجرودة جديدة لدمجها في هذا الملف."
            status_text.color = "#1B2A47"
            
    except Exception as ex:
        status_text.value = f"❌ خطأ أثناء دمج وتجميع الملف: {str(ex)}"
        status_text.color = "#850000"
        
    page.update()

def handle_employee_change(e):
    """دالة الاستدعاء لتغيير اسم الموظف وتحديث الهوية من النافذة الجانبية للسهم"""
    global current_employee
    new_name = str(change_name_input.value).strip()
    if new_name:
        current_employee = new_name
        employee_label.value = f"👤 الموظف المسؤول الحالي: {current_employee}"
        employee_label.color = "#1B2A47"
        employee_sidebar_text.value = f"الموظف الحالي: {current_employee}"
        status_text.value = f"✅ تم تحديث وتغيير اسم الموظف بنجاح إلى: {current_employee}"
        status_text.color = "#008000"
        change_name_input.value = ""
        try:
            page.drawer.open = False
        except:
            pass
    page.update()

def handle_file_upload_simulation(e):
    """قراءة وتثبيت ملف النظام المرجعي الأساسي مع حماية البيانات المخزنة من التصفير تلقائياً"""
    global master_df, current_jard_df
    
    # 1. تنشيط واستدعاء قواعد البيانات المحلية من القرص أولاً
    init_offline_database()
    
    if not master_df.empty:
        # 2. إذا كان كاش الجرد المحلي فارغاً (أول تشغيل)، نقوم بتأسيسه بالكامل بناءً على الماستر داتا
        if current_jard_df.empty:
            current_jard_df = pd.DataFrame()
            current_jard_df['رمز المادة'] = master_df['رمز المادة'].astype(str).apply(clean_value)
            current_jard_df['الرمز / الباركود'] = master_df['الباركود'].astype(str).apply(clean_value)
            current_jard_df['اللون'] = master_df['اللون'].astype(str).apply(clean_value)
            current_jard_df['القياس'] = master_df['القياس'].astype(str).apply(clean_value)
            current_jard_df['الأصلية'] = master_df['الكمية_الأصلية'].astype(str).apply(clean_value)
            current_jard_df['المجرودة'] = "0"
            
            for idx, row in current_jard_df.iterrows():
                orig_val = int(float(row['الأصلية'])) if row['الأصلية'] else 0
                current_jard_df.at[idx, 'الفرق'] = str(-orig_val)
        else:
            # 3. [حماية وحقن البيانات] إذا كان هناك جرد مستمر، نقوم فقط بتحديث كميات الماستر داتا المرجعية
            # دون مساس بالكميات المجرودة التي عَدّها الموظف سابقاً لكي يكمل عمل اليوم التالي بأمان
            for idx, row in master_df.iterrows():
                m_barcode = str(row.get('الباركود', '')).strip().upper()
                m_color = str(row.get('اللون', '')).strip()
                m_size = str(row.get('القياس', '')).strip()
                m_orig = int(float(row.get('الكمية_الأصلية', 0))) if row.get('الكمية_الأصلية', '') else 0
                
                # البحث عن المادة في السجل النشط الحالي لتحديث قيمتها الأصلية
                match_idx = current_jard_df[
                    (current_jard_df['الرمز / الباركود'] == m_barcode) & 
                    (current_jard_df['اللون'] == m_color) & 
                    (current_jard_df['القياس'] == m_size)
                ].index
                
                if len(match_idx) > 0:
                    c_idx = match_idx[0]
                    current_jard_df.at[c_idx, 'الأصلية'] = str(m_orig)
                    try:
                        c_jard = int(float(current_jard_df.at[c_idx, 'المجرودة']))
                    except:
                        c_jard = 0
                    current_jard_df.at[c_idx, 'الفرق'] = str(c_jard - m_orig)
                    
        # 4. حفظ الكاش المحدث على القرص الصلب لحمايته عند الإغلاق
        try:
            current_jard_df.to_excel(LOCAL_JARD_FILE, index=False)
        except:
            pass
            
        status_text.value = "💙 تم قراءة الماستر داتا بنجاح وتأمين جرد الأيام السابقة دون فقدان أي قطعة!"
        status_text.color = "#1B2A47"
        
        # فتح صندوق تسجيل اسم الموظف الحالي الميداني
        open_name_registration_dialog()
    else:
        status_text.value = "⚠️ لم يتم العثور على ملف master_data_reference.xlsx بجانب التطبيق أوفلاين!"
        status_text.color = "#850000"
    page.update()

# 1. الإعلان المسبق عن كافة الكائنات البرمجية لمنع التداخل والتعارض البصري
status_text = ft.Text("● النظام نشط ومستعد أوفلاين... يرجى رفع ملف النظام المرجعي أولاً", size=13, weight=ft.FontWeight.W_500, color="#1B2A47")
employee_label = ft.Text("⚠️ يرجى تسجيل اسم الموظف أولاً", size=14, color="#850000", weight=ft.FontWeight.BOLD)
employee_sidebar_text = ft.Text("الموظف الحالي: لم يسجل بعد", size=14, color="#FFFFFF")

laser_input = ft.TextField(label="📷 امسح بالليزر مباشرة (Focus المستمر):", border_color="#1B2A47", width=float("inf"), text_align=ft.TextAlign.RIGHT, rtl=True, disabled=True)
manual_input = ft.TextField(label="✍️ الكتابة اليدوية للكود:", border_color="#1B2A47", width=float("inf"), text_align=ft.TextAlign.RIGHT, rtl=True, disabled=True)
camera_checkbox = ft.Checkbox(label="📸 تفعيل كاميرا الجوال كقارئ باركود احتياطي", value=False, disabled=True) # [إضافة التعطيل هنا]

view_mode_radio = ft.RadioGroup(
    content=ft.Row([
        ft.Radio(value="all", label="عرض الجدول كاملاً"),
        ft.Radio(value="last", label="عرض السطر الأخير المقروء فقط")
    ], alignment=ft.MainAxisAlignment.START, spacing=20),
    value="all"
)

search_field = ft.TextField(label="🔍 خانة البحث السريع (رمز، باركود، لون):", border_color="#1B2A47", width=float("inf"), text_align=ft.TextAlign.RIGHT, rtl=True)
edit_barcode_field = ft.TextField(label="🧩 كود المادة للتصحيح:", border_color="#1B2A47")
edit_qty_field = ft.TextField(label="🔢 الكمية الجديدة:", value="0", border_color="#1B2A47")
change_name_input = ft.TextField(label="👤 اكتب الاسم الجديد هنا:", border_color="#FFFFFF", label_style=ft.TextStyle(color="#FFFFFF"), text_style=ft.TextStyle(color="#FFFFFF"))

jard_table = ft.DataTable(
    columns=[
        ft.DataColumn(ft.Text("رمز المادة", weight=ft.FontWeight.BOLD)),
        ft.DataColumn(ft.Text("الرمز / الباركود", weight=ft.FontWeight.BOLD)),
        ft.DataColumn(ft.Text("اللون", weight=ft.FontWeight.BOLD)),
        ft.DataColumn(ft.Text("القياس", weight=ft.FontWeight.BOLD)),
        ft.DataColumn(ft.Text("الأصلية", weight=ft.FontWeight.BOLD)),
        ft.DataColumn(ft.Text("المجرودة", weight=ft.FontWeight.BOLD)),
        ft.DataColumn(ft.Text("الفرق", weight=ft.FontWeight.BOLD)),
    ],
    rows=[]
)

name_dialog_input = ft.TextField(label="اكتب اسمك الثلاثي الصريح:", border_color="#1B2A47", autofocus=True)

name_dialog_box = ft.AlertDialog(
    modal=True,
    title=ft.Row([ft.Icon("person", color="#1B2A47"), ft.Text("تسجيل هوية الموظف الميداني", weight=ft.FontWeight.BOLD)], spacing=10),
    content=ft.Container(
        content=ft.Column([
            ft.Container(content=name_dialog_input, width=320, height=65),
            ft.Row([
                ft.TextButton(
                    content=ft.Container(
                        content=ft.Text("🎯 تفعيل وحفظ الهوية", color="#FFFFFF", weight=ft.FontWeight.BOLD),
                        alignment="center",
                        bgcolor="#1B2A47",
                        height=42,
                        width=160,
                        border_radius=8
                    ),
                    on_click=lambda e: submit_initial_name(e)
                )
            ], alignment=ft.MainAxisAlignment.END)
        ], tight=True, spacing=10),
        width=340,
        height=140
    )
)

def open_name_registration_dialog():
    """تخطي الصندوق المنبثق المسبب للفصل والتفعيل التلقائي لواجهة التطبيق"""
    global current_employee
    current_employee = "الموظف المسؤول"
    if 'employee_label' in globals():
        employee_label.value = f"👤 الموظف المسؤول الحالي: {current_employee}"
    page.update()

def submit_initial_name(e):
    """حفظ الاسم الثلاثي الأولي وتفعيل الهوية وفك حظر حقول الإدخال والكاميرا للجرد"""
    global current_employee
    name_val = str(name_dialog_input.value).strip()
    if name_val:
        current_employee = name_val
        employee_label.value = f"👤 الموظف المسؤول الحالي: {current_employee}"
        employee_label.color = "#1B2A47"
        employee_sidebar_text.value = f"الموظف الحالي: {current_employee}"
        
        # [فك الحظر الشامل فوراً] تفعيل حقول الإدخال ومربع الكاميرا الاحتياطية
        laser_input.disabled = False
        manual_input.disabled = False
        camera_checkbox.disabled = False
        
        name_dialog_box.open = False
        status_text.value = "✅ تم تفعيل الملف وتسجيل الموظف بنجاح! النظام مستعد تماماً للجرد."
        status_text.color = "#008000"
        rebuild_jard_table_rows()
        laser_input.focus()
    page.update()

def handle_employee_change(e):
    """تحديث اسم الموظف من السهم الجانبي وتأكيد تفعيل الحقول والكاميرا"""
    global current_employee
    new_name = str(change_name_input.value).strip()
    if new_name:
        current_employee = new_name
        employee_label.value = f"👤 الموظف المسؤول الحالي: {current_employee}"
        employee_label.color = "#1B2A47"
        employee_sidebar_text.value = f"الموظف الحالي: {current_employee}"
        
        # تأكيد فتح الحقول والكاميرا في حال تم التغيير من السهم
        laser_input.disabled = False
        manual_input.disabled = False
        camera_checkbox.disabled = False
        
        status_text.value = f"✅ تم تحديث وتغيير اسم الموظف بنجاح إلى: {current_employee}"
        status_text.color = "#008000"
        change_name_input.value = ""
        try:
            page.drawer.open = False
        except:
            pass
    page.update()

def trigger_wipe_and_reset(e):
    """تصفير وإلغاء جرد الموظف الحالي بالكامل والبدء من جديد دون مساس بالآخرين"""
    global current_jard_df, LOCAL_JARD_FILE
    
    if not current_jard_df.empty:
        current_jard_df['المجرودة'] = "0"
        for idx, row in current_jard_df.iterrows():
            orig_val = int(float(row['الأصلية'])) if row['الأصلية'] else 0
            current_jard_df.at[idx, 'الفرق'] = str(-orig_val)
        
        try:
            current_jard_df.to_excel(LOCAL_JARD_FILE, index=False)
        except:
            pass
            
    emp_name = current_employee if current_employee else "عام"
    emp_file = os.path.join(BASE_PATH, f"scans_{emp_name}.xlsx")
    if os.path.exists(emp_file):
        try:
            os.remove(emp_file)
        except:
            pass
            
    status_text.value = f"🗑️ تم مسح وتصفير سجلات جرد الموظف ({emp_name}) بالكامل والبدء من جديد!"
    status_text.color = "#850000"
    rebuild_jard_table_rows()
    page.update()


def main(page_obj):
    global page, view_mode_radio, BASE_PATH, MASTER_FILE, LOCAL_JARD_FILE
    # ... بقية كود دالة ماعين بالأسفل
   
    # تحديد مجلد التخزين الآمن والمسموح بالكتابة فيه تلقائياً على الأندرويد
    BASE_PATH = page_obj.user_data_dir
    MASTER_FILE = os.path.join(BASE_PATH, "master_data_reference.xlsx")
    LOCAL_JARD_FILE = os.path.join(BASE_PATH, "local_employee_jard.xlsx")
    
    # تفعيل حظر الإغلاق التلقائي للشاشة ميدانياً أثناء الجرد المستمر للأندرويد
    page_obj.keep_awake = True

    # توجيه الدالة لاستخدام متغير الموظف العام الموحد
    global current_employee
    
    # 1. تعريف دوال معالجة الأحداث الميدانية محلياً لحمايتها من الـ NameError
    def handle_laser_scan(e):
        barcode = str(laser_input.value).strip()
        if barcode:
            process_barcode_scan(barcode)
        laser_input.value = ""  
        laser_input.focus()    
        page.update()

    def handle_manual_submit(e):
        barcode = str(manual_input.value).strip()
        if barcode:
            process_barcode_scan(barcode)
        manual_input.value = ""  
        page.update()

    # تعريف الكائن الخاص بالتقاط وتجميع ملفات التقارير للأندرويد
    global sync_file_picker
    sync_file_picker = ft.FilePicker(on_result=handle_sync_file_result)
    
    # حقن الكائن في الصفحة ليصبح متاحاً للاستدعاء بلحظتها
    page_obj.overlay.append(sync_file_picker)

    page = page_obj
    page.title = "🔦 نظام جرد Torino الذكي"
    page.rtl = True
    page.scroll = ft.ScrollMode.AUTO
    page.theme_mode = ft.ThemeMode.LIGHT
    
    # 2. ربط الأحداث الميدانية اللحظية بالدوال الداخلية للسرعة الفائقة والتفريغ
    laser_input.on_submit = handle_laser_scan
    manual_input.on_submit = handle_manual_submit
    search_field.on_change = lambda _: rebuild_jard_table_rows()
    view_mode_radio.on_change = lambda _: rebuild_jard_table_rows()
 
    # هندسة أبعاد الارتفاعات لتتطابق مع شاشات الأندرويد والهاتف
    laser_input.height = 55
    manual_input.height = 55
    search_field.height = 55
    edit_barcode_field.height = 50
    edit_qty_field.height = 50

    # 1. تفعيل القفل الحديدي الأولي على الحقول لحظر التعبئة قبل إثبات الهوية
    laser_input.disabled = True
    manual_input.disabled = True
    search_field.disabled = True

    # 2. دالة رفع الحظر والاعتماد الفوري للفوكس الليزري المستمر
    def submit_employee_identity(e):
        global current_employee
        entered_name = employee_name_field.value.strip()
        
        if not entered_name:
            employee_name_field.error_text = "الاسم الثلاثي الصريح إجباري لبدء المنظومة!"
            page_obj.update()
            return
        
        # حفظ الهوية في المتغير المشترك وتثبيتها بالكاش
        current_employee = entered_name
        identity_dialog.open = False  # إغلاق النافذة الإجبارية
        
        # رفع الحظر الحديدي عن الحقول الميدانية بلحظتها
        laser_input.disabled = False
        manual_input.disabled = False
        search_field.disabled = False
        
        # توجيه سرعة الفوكس المستمر والكامل لليزر فوراً لمنع التشتت
        laser_input.focus()
        
        page_obj.show_snack_bar(
            ft.SnackBar(ft.Text(f"⚙️ تم تفعيل منظومة Torino للموظف: {current_employee}"), bgcolor=ft.colors.GREEN_700)
        )
        page_obj.update()

    # 3. بناء نافذة الـ AlertDialog الإجبارية المخصصة أوفلاين
    employee_name_field = ft.TextField(
        label="الاسم الثلاثي الصريح للموظف المسؤول", 
        autofocus=True,
        on_submit=submit_employee_identity
    )
    
    identity_dialog = ft.AlertDialog(
        modal=True,  
        title=ft.Text("🔒 قفل أمان الهوية والتأمين الميداني"),
        content=ft.Column([
            ft.Text("يرجى إدخال اسمك الثلاثي لفتح صلاحيات حقول الجرد..."),
            employee_name_field
        ], tight=True, spacing=15),
        actions=[
            ft.ElevatedButton("حفظ، اعتماد وتفعيل الفوكس 🚀", on_click=submit_employee_identity)
        ]
    )
    
    # 4. تضع الأسطر الثلاثة هنا مباشرة (تحت تعريف identity_dialog) وبنفس المستوى:
    page_obj.dialog = identity_dialog
    identity_dialog.open = True
    page_obj.update()

    # 2. بناء وتصميم الدرج الجانبي (Drawer) المخصص للسهم لتغيير اسم الموظف بأي وقت
    page.drawer = ft.NavigationDrawer(
        controls=[
            ft.Container(
                content=ft.Column([
                    ft.Icon(ft.icons.ACCOUNT_CIRCLE, size=50, color="#FFFFFF"),
                    employee_sidebar_text,
                    ft.Divider(color="#FFFFFF", thickness=1),
                    change_name_input,
                    ft.ElevatedButton("💾 حفظ وتحديث الهوية", on_click=handle_employee_change, bgcolor="#FFFFFF", color="#1B2A47")
                ], spacing=15),
                bgcolor="#1B2A47", padding=20, border_radius=ft.border_radius.only(top_left=16, bottom_left=16)
            )
        ]
    )
    
    # 3. حقن وهندسة التصميم النهائي والكامل لتطبيق الهاتف للأندرويد
    page.add(
        ft.Row([
            ft.IconButton(ft.icons.MENU, icon_color="#1B2A47", on_click=lambda _: setattr(page.drawer, 'open', True) or page.update()),
            ft.Text("🔦 نظام جرد Torino الذكي", size=18, weight=ft.FontWeight.BOLD, color="#1B2A47")
        ], alignment=ft.MainAxisAlignment.START),
        
        ft.Container(content=status_text, bgcolor="#FFF9E6", padding=10, border_radius=8, margin=ft.margin.only(bottom=5)),
        
        ft.Container(
            content=ft.ElevatedButton(
                "📥 اضغط لتثبيت وقراءة ملف النظام المرجعي", 
                icon=ft.icons.UPLOAD_FILE, bgcolor="#1B2A47", color="#FFFFFF", 
                on_click=handle_file_upload_simulation, height=50
            ),
            width=float("inf")
        ),
        
        ft.Divider(color="#D4AF37", thickness=2, height=20),
        employee_label,
        
        ft.Column([
            ft.Text("🔍 مركز المسح والإدخال الميداني:", size=14, weight=ft.FontWeight.BOLD, color="#1B2A47"),
            ft.Container(content=laser_input, width=float("inf")),
            ft.Container(content=manual_input, width=float("inf")),
        ], spacing=10),
        
        ft.Container(content=camera_checkbox, padding=ft.padding.only(top=5, bottom=5)),
        ft.Divider(color="#1B2A47", thickness=1),
        
        ft.Column([
            ft.Row([
                ft.Icon(ft.icons.GRID_VIEW, color="#1B2A47", size=20),
                ft.Text("وضع العرض المفضل للجدول:", size=13, weight=ft.FontWeight.BOLD, color="#1B2A47")
            ], spacing=5),
            view_mode_radio
        ], spacing=5),
        
        ft.Container(padding=ft.padding.only(bottom=5)),
        ft.Text("📋 جدول كميات الجرد الحالية ومطابقة الألوان:", size=14, weight=ft.FontWeight.BOLD, color="#1B2A47"),
        ft.Container(content=jard_table, padding=5, border_radius=5, bgcolor="#F5F5F5"),
        
        ft.Divider(color="#1B2A47", thickness=1, height=25),
        
        ft.Column([
            ft.Text("✏️ تعديل الكميات والبحث السريع:", size=14, weight=ft.FontWeight.BOLD, color="#1B2A47"),
            ft.Container(content=search_field, width=float("inf")),
            ft.Row([
                ft.Container(content=edit_barcode_field, expand=2),
                ft.Container(content=edit_qty_field, expand=1),
            ], spacing=5),
            ft.ElevatedButton("✅ تطبيق تعديل الكمية المحددة", on_click=apply_quantity_correction, bgcolor="#008000", color="#FFFFFF", height=45, width=float("inf"))
        ], spacing=10),
        
        ft.Divider(color="#1B2A47", thickness=1),
        
        ft.Column([
            ft.ElevatedButton("⚡ تفعيل المزامنة ودمج الجرد أوفلاين", icon=ft.icons.SYNC, bgcolor="#1B2A47", color="#FFFFFF", on_click=trigger_manual_sync, height=50, width=float("inf")),
            ft.ElevatedButton("📥 تحميل تقرير الجرد الختامي المجمع", icon=ft.icons.DOWNLOAD, bgcolor="#008000", color="#FFFFFF", on_click=trigger_export_excel, height=50, width=float("inf")),
            ft.ElevatedButton("🗑️ مسح الجرد والبدء من جديد", icon="delete_forever", bgcolor="#850000", color="#FFFFFF", on_click=trigger_wipe_and_reset, height=50, width=float("inf")),
        ], spacing=12)
    )
    
    # [تحديث الجزء الأخير في دالة main قبل page.update]
    init_offline_database()
    if not current_jard_df.empty:
        status_text.value = "📥 تم استعادة بيانات الجرد السابق تلقائياً أوفلاين! يرجى تسجيل الاسم لإكمال العمل."
        status_text.color = "#008000"
        # إجبار فتح صندوق الاسم فوراً لحماية النظام من النسيان
        open_name_registration_dialog()
        employee_label.value = "👤 تم استعادة الجرد النشط، يرجى إدخال اسمك عبر السهم الجانبي عند الحاجة للتوثيق"
        employee_label.color = "#1B2A47"
    
    rebuild_jard_table_rows()
    page.update()

if __name__ == "__main__":
    ft.app(target=main)
