import json
import os
import sys
import requests
import pandas as pd
from tqdm import tqdm # For tracking progress across many files
import re # لاستخدام التعبيرات العادية لتنظيف محتوى JSON
import time # لاستخدام التأخير الزمني في إعادة المحاولة (ولتحويل الـ timestamp)

# ==============================================================================
#                 **CONFIGURATION / الإعدادات (هام جداً للمراجعة!)**
# ==============================================================================

API_KEY = os.getenv("API_KEY")



# 2. نقطة نهاية API الجديدة
OCR_API_URL = "https://extraction-api.nanonets.com/extract" 
# OCR_API_URL = "https://extraction-api.nanonets.com/api/v1/extract/sync"

# 3. Image Source Directory
IMAGE_ROOT_DIR = r"C:\Users\t\Downloads\docstrange-main\docstrange-main\carr_ocr\flyers_highres"

# 🛑 4. **نقطة البدء الجديدة (للتشغيل الجزئي)** #    - أدخل اسم الملف الذي تود البدء منه (مثل 'page_72.jpg') لتخطي الملفات التي تمت معالجتها.
#    - اتركها None للبدء من أول ملف (إذا كنت تريد إعادة معالجة الكل).
# START_FILE_NAME = "page_1.jpg" # ابدأ من جديد
START_FILE_NAME="20251218_F1_flyer_5897_page_1.jpg"
# 5. JSON Schema for Extraction (مخطط الاستخراج - تم إضافة جميع الحقول الجديدة)
JSON_SCHEMA = {
    "type": "object",
    "description": "Information extracted from the flyer, including the supermarket name and a list of all products.",
    "properties": {
        "supermarket_name": {
            "type": "string",
            "description": "اسم السوبر ماركت أو المتجر الرئيسي الذي أصدر الفلاير."
        },
        "offer_validity_period": { # 👈 حقل جديد على مستوى الفلاير
            "type": "string",
            "description": "فترة سريان العروض في الفلاير بالكامل (مثلاً 2025-12-16 to 2025-12-31)."
        },
        "location_branch": { # 👈 حقل جديد على مستوى الفلاير
            "type": "string",
            "description": "اسم الفرع أو المدينة المذكورة في الإعلان (إذا لم يكن العرض عاماً)."
        },
        "products": {
            "type": "array",
            "description": "List of grocery, household, and food products with their unit, price, and discount information.",
            "items": {
                "type": "object",
                "properties": {
                    "category": {"type": "string", "description": "الفئة التي ينتمي إليها المنتج"},
                    "name": {"type": "string", "description": "اسم المنتج الأساسي باللغة اللاتينية/الإنجليزية"},
                    "arabic_name": {"type": "string", "description": "الاسم العربي للمنتج، بالضبط كما يظهر في الصورة."},
                    "brand_name": {"type": "string", "description": "الاسم التجاري أو العلامة التجارية للمنتج."},
                    "unit": {"type": "string", "description": "وحدة القياس (مثل: كجم، لتر، عبوة)."},
                    "original_price": {"type": "number", "description": "السعر قبل الخصم، قيمة رقمية."},
                    "discounted_price": {"type": "number", "description": "السعر الحالي بعد الخصم، قيمة رقمية."},
                    "discount_percentage": {"type": "integer", "description": "نسبة الخصم المُعلنة كقيمة عدد صحيح."},
                    "price_currency": {"type": "string", "description": "العملة المستخدمة في التسعير (مثل SAR)."}, # 👈 حقل جديد
                    "constraints_limits": {"type": "string", "description": "أي قيود أو حدود للشراء (مثل: 'الحد الأقصى 2 لكل عميل')."}, # 👈 حقل جديد
                    "bounding_box" : {"type": "string", "description": "الاحداثيات"},
                    "confidence_score": {"type": "string", "description": "الموثوقية"},
                    "promotional_tagline": {"type": "string", "description": "العبارة الترويجية أو الوصف القصير المكتوب بجانب المنتج."} # 👈 حقل جديد
                },
                "required": ["name", "unit"]
            }
        }
    },
    "required": ["supermarket_name", "products"] 
}
SCHEMA_JSON_STRING = json.dumps(JSON_SCHEMA)

# 6. Output Excel File
OUTPUT_EXCEL_FILENAME = "./output/aggregated_bulk_flyer_data_unified.xlsx" 

# 7. إعدادات إعادة المحاولة
MAX_RETRIES = 3 
INITIAL_BACKOFF = 5 

# ==============================================================================

def get_mime_type(file_path):
    """يحدد نوع MIME للملف بناءً على امتداده."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.jpg', '.jpeg']:
        return 'image/jpeg'
    elif ext == '.png':
        return 'image/png'
    return 'application/octet-stream' 

def clean_and_parse_json(json_string, file_name):
    """نسخة مطورة لتحليل البيانات حتى لو جاءت بتنسيق مختلف"""
    if not json_string:
        return {}
    
    # تنظيف النصوص الزائدة التي قد يضيفها Gemini مثل ```json
    json_string = re.sub(r'```json\s?|```', '', json_string).strip()
    
    match = re.search(r'\{.*\}', json_string, re.DOTALL)
    clean_json_str = match.group(0) if match else json_string

    try:
        data = json.loads(clean_json_str)
        # التحقق من وجود المنتجات بصرف النظر عن هيكل الرد
        if isinstance(data, dict):
            if 'products' in data:
                return data
            # إذا كان الرد يحتوي على مفتاح آخر بداخله المنتجات
            for key in data:
                if isinstance(data[key], dict) and 'products' in data[key]:
                    return data[key]
        return {}
    except Exception as e:
        tqdm.write(f"   ❌ فشل تحليل JSON لـ {file_name}: {e}")
        # طباعة جزء من الرد لمعرفة المشكلة
        tqdm.write(f"   📝 بداية الرد المستلم: {json_string[:100]}...") 
        return {}

def _process_image_no_retry(file_path):
    if not API_KEY:
        raise Exception("API_KEY is not set.")

    HEADERS = {'Authorization': f'Bearer {API_KEY}'}
    file_mime_type = get_mime_type(file_path)

    with open(file_path, 'rb') as f:
        files = {'file': (os.path.basename(file_path), f, file_mime_type)}
        data = {
            'output_type': 'specified-json',
            'json_schema': SCHEMA_JSON_STRING,
            'model_type': 'nanonets-ocr-3' 
        }

        response = requests.post(OCR_API_URL, headers=HEADERS, files=files, data=data, timeout=90)
    
    if response.status_code != 200:
        tqdm.write(f"   🚫 خطأ من الخادم ({response.status_code}): {response.text}")
        response.raise_for_status()

    api_result = response.json()
    
    # تحسين: البحث عن المحتوى في أكثر من مكان (content أو result)
    extracted_content_str = api_result.get("content") or api_result.get("result", [{}])[0].get("prediction", "")
    
    if not extracted_content_str:
        # إذا فشل كل شيء، اطبع هيكل الرد كاملاً للتشخيص
        tqdm.write(f"   ❓ رد غير متوقع من API: {api_result}")
        return None

    return clean_and_parse_json(str(extracted_content_str), os.path.basename(file_path))

def process_image_with_retries(file_path):
    """
    يقوم بمعالجة صورة مع آلية إعادة المحاولة للأخطاء العابرة (شبكة، مهلة، JSON مقطوع).
    """
    file_name = os.path.basename(file_path)
    
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return _process_image_no_retry(file_path)

        except (requests.exceptions.RequestException, ValueError) as e:
            
            if isinstance(e, requests.exceptions.HTTPError) and e.response.status_code in [400, 401]:
                if e.response.status_code == 401:
                    tqdm.write(f"   ❌ خطأ 401 (غير مصرح): المفتاح غير صالح. (المحاولة {attempt}/{MAX_RETRIES})")
                else:
                    tqdm.write(f"   ❌ خطأ 400 (طلب غير صحيح): توقف عن المحاولة. (المحاولة {attempt}/{MAX_RETRIES})")
                return None
            
            if attempt < MAX_RETRIES:
                wait_time = INITIAL_BACKOFF * (2 ** (attempt - 1))
                error_type = "JSON Decoded/لا يوجد محتوى" if isinstance(e, ValueError) else "شبكة/مهلة"
                tqdm.write(f"   ⚠️ خطأ {error_type} لـ {file_name}: {e.__class__.__name__}. جاري الانتظار {wait_time} ثوان وإعادة المحاولة... ({attempt}/{MAX_RETRIES})")
                time.sleep(wait_time)
            else:
                tqdm.write(f"   ❌ فشلت المعالجة لـ {file_name} بعد {MAX_RETRIES} محاولات.")
                return None
        
        except Exception as e:
            tqdm.write(f"   ❌ خطأ غير متوقع أثناء معالجة {file_name}: {e}")
            return None
            
    return None

def find_all_images(root_dir):
    """يتجول في الدليل الجذري ويجد جميع ملفات الصور (jpg, jpeg, png)."""
    image_paths = []
    
    for dirpath, dirnames, filenames in os.walk(root_dir):
        for file in filenames:
            if file.lower().endswith(('.jpg', '.jpeg', '.png')):
                image_paths.append(os.path.join(dirpath, file))
    
    return image_paths

def export_to_excel(data, filename):
    """يحول القائمة المجمعة للبيانات إلى DataFrame ويصدرها كملف Excel."""
    print("\n=========================================================")
    print("💾 جاري حفظ البيانات المجمعة...")
    
    df = pd.DataFrame(data)

    def calculate_discounted_price(row):
        try:
            original = float(row['original_price'])
            discount = float(row['discount_percentage'])
            if original and discount:
                return original * (1 - discount / 100)
        except (ValueError, TypeError, KeyError):
            return None
            
    df['calculated_discounted_price'] = df.apply(calculate_discounted_price, axis=1)

    # 💡 خريطة الأعمدة المُحدَّثة مع جميع الإضافات الجديدة
    col_map = {
        "supermarket_name": "اسم السوبر ماركت", 
        "offer_validity_period": "فترة سريان العرض", # 👈 جديد
        "location_branch": "الفرع / الموقع", # 👈 جديد
        "file_modified_date": "تاريخ تعديل ملف الفلاير", 
        "category": "الفئة", 
        "brand_name": "اسم العلامة التجارية", 
        "name": "اسم المنتج (الأساسي)", 
        "arabic_name": "الاسم بالعربية",
        "promotional_tagline": "العبارة الترويجية", # 👈 جديد
        "unit": "الوحدة",
        "original_price": "السعر الأصلي", 
        "discounted_price": "السعر الحالي",
        "discount_percentage": "الخصم %", 
        "price_currency": "العملة", # 👈 جديد
        "constraints_limits": "قيود العرض", # 👈 جديد
        "source_file": "اسم ملف الفلاير", 
        "subfolder_name": "اسم المجلد الفرعي", 
        "bounding_box": "الاحداثيات",
        "confidence_score" : "الموثوقية",
        "calculated_discounted_price": "السعر المحسوب"
    }
    
    df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)
    
    final_cols = [c for c in col_map.values() if c in df.columns]
    df = df[final_cols]

    try:
        df.to_excel(filename, index=False, engine='xlsxwriter')
        print(f"✅ تم الحفظ بنجاح ({len(df)} صف) في: {filename}")
    except PermissionError:
        print(f"❌ خطأ في الإذن: لا يمكن الكتابة إلى الملف '{filename}'.")
        print("   الرجاء التأكد من إغلاق الملف قبل المحاولة مرة أخرى.")
    except Exception as e:
        print(f"❌ فشل التصدير إلى Excel: {e}")


def main():
    if not os.path.exists(IMAGE_ROOT_DIR):
        print(f"❌ خطأ: المجلد الجذري للصور غير موجود: {IMAGE_ROOT_DIR}")
        return

    all_image_paths = find_all_images(IMAGE_ROOT_DIR)
    
    if not all_image_paths:
        print(f"❌ لم يتم العثور على أي ملفات صور (jpg/jpeg/png) في: {IMAGE_ROOT_DIR}")
        return

    # 🛑 1. تطبيق منطق الاستئناف
    start_index = 0
    if START_FILE_NAME and START_FILE_NAME != "None":
        try:
            full_start_path = next(p for p in all_image_paths if os.path.basename(p) == START_FILE_NAME)
            start_index = all_image_paths.index(full_start_path)
            remaining_image_paths = all_image_paths[start_index:]
            tqdm.write(f"**استئناف المعالجة:** سيتم البدء من الملف {START_FILE_NAME} (تخطي {start_index} ملف).")
        except StopIteration:
            tqdm.write(f"❌ خطأ: الملف المحدد للبدء ({START_FILE_NAME}) غير موجود في المجلد.")
            return
    else:
        remaining_image_paths = all_image_paths 

    total_files = len(all_image_paths) 
    
    print(f"\n=========================================================")
    print(f"💡 تم العثور على {len(remaining_image_paths)} صورة فلاير للمعالجة (من أصل {total_files}).")
    print(f"المسار الجذري للمجلد: {IMAGE_ROOT_DIR}")
    print("=========================================================")

    all_extracted_data = []
    
    for file_path in tqdm(remaining_image_paths, desc="معالجة الصور", initial=start_index, total=total_files):
        
        # حساب تاريخ التعديل لملف الفلاير نفسه
        try:
            mtime_timestamp = os.path.getmtime(file_path)
            # تنسيق التاريخ والوقت (YYYY-MM-DD HH:MM:SS)
            file_modified_date = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(mtime_timestamp))
        except OSError:
            file_modified_date = 'غير متوفر'
            
        # حساب المسار النسبي والمجلد الفرعي
        flyer_relative_path = os.path.relpath(file_path, IMAGE_ROOT_DIR)
        subfolder_name = os.path.dirname(flyer_relative_path)
        if subfolder_name == '.': 
            subfolder_name = 'Root Directory'
            
        extracted_data = process_image_with_retries(file_path)
        
        if not extracted_data or not extracted_data.get('products'):
            continue

        if extracted_data and extracted_data.get('products'):
            
            flyer_supermarket_name = extracted_data.get('supermarket_name', 'غير محدد')
            
            # استخراج الحقول الجديدة على مستوى الفلاير
            offer_validity_period = extracted_data.get('offer_validity_period', 'غير مستخرج')
            location_branch = extracted_data.get('location_branch', 'غير مستخرج')
            
            for product in extracted_data['products']:
                if isinstance(product, dict):
                    
                    product['source_file'] = os.path.basename(file_path)
                    product['supermarket_name'] = flyer_supermarket_name 
                    product['subfolder_name'] = subfolder_name         
                    product['file_modified_date'] = file_modified_date 
                    
                    # ربط الحقول الجديدة على مستوى الفلاير بكل منتج
                    product['offer_validity_period'] = offer_validity_period
                    product['location_branch'] = location_branch

                    # تنظيف حقول غير مرغوبة
                    if 'product_image_link' in product:
                        del product['product_image_link']
                        
                    all_extracted_data.append(product)
        
    if all_extracted_data:
        if START_FILE_NAME and START_FILE_NAME != "None" and start_index > 0:
            temp_filename = "./output/aggregated_bulk_flyer_data_NEW_RUN.xlsx"
            print("\n=========================================================")
            print("⚠️ ملاحظة: تم إنشاء ملف جديد للبيانات المُستأنفة فقط لتجنب تكرار البيانات السابقة.")
            export_to_excel(all_extracted_data, temp_filename)
            print(f"يرجى دمج محتوى {temp_filename} مع ملفك الأصلي ({OUTPUT_EXCEL_FILENAME}).")
        else:
            export_to_excel(all_extracted_data, OUTPUT_EXCEL_FILENAME)
    else:
        print("\n❌ لم يتم استخراج أي بيانات منتج بنجاح من جميع الصور.")


if __name__ == "__main__":
    try:
        # التأكد من وجود التبعيات الضرورية
        import pandas
        import openpyxl 
        import requests
        from tqdm import tqdm
        import xlsxwriter
        
    except ImportError as e:
        print(f"يرجى تثبيت المكتبات المطلوبة: pip install pandas openpyxl requests tqdm xlsxwriter")
        print(f"الخطأ المفقود: {e}")
        sys.exit(1)
        
    try:
            main()
    except KeyboardInterrupt:
        print("\nتم إيقاف العملية بواسطة المستخدم.")
