import requests
import sqlite3
import time
import json

# config.json faylidan API ma'lumotlarini yuklash
with open("config.json", "r") as file:
    config = json.load(file)

BASE_URL = f"{config['student_otm_url']}rest/v1/data/student-list"  # API manzili (taxminiy, to'g'rilang)
api_key = config["api_key"]
HEADERS = {"Authorization": f"Bearer {api_key}"}

# SQLite bazaga ulanish
conn = sqlite3.connect("database/hemis_api.db")
cursor = conn.cursor()

# Jadval yaratish
cursor.execute('''
CREATE TABLE IF NOT EXISTS student_list (
    id INTEGER PRIMARY KEY,
    full_name TEXT,
    third_name TEXT,
    student_id_number TEXT,
    department_id INTEGER,
    department_name TEXT,
    group_id INTEGER,
    group_name TEXT,
    level_code TEXT,
    level_name TEXT,
    education_year_code TEXT,
    education_year_name TEXT,
)
''')
conn.commit()

def fetch_data(page):
    """API'dan ma'lumot olish"""
    params = {"page": page, "limit": 200, "_student_status":-1}
    response = requests.get(BASE_URL, headers=HEADERS, params=params)
    if response.status_code == 200:
        return response.json()
    return None

def fetch_and_store_students():
    """Ma'lumotlarni API'dan yuklab, SQLite bazasiga saqlash"""
    page = 1
    data = fetch_data(page)
    
    if not data:
        print("❌ Ma'lumotni olishda xatolik yuz berdi.")
        return
    
    # Pagination ma'lumotlarini xavfsiz olish
    try:
        data_content = data.get("data", {})
        pagination = data_content.get("pagination", {})
        page_count = pagination.get("pageCount", 1)
    except (IndexError, KeyError) as e:
        print(f"❌ Pagination ma'lumotlarini olishda xato: {e}")
        page_count = 1
    
    print(f"📌 Jami sahifalar soni: {page_count}")
    
    while page <= page_count:
        print(f"🔄 {page} - ma'lumot olinmoqda...")
        data = fetch_data(page)
        
        if not data:
            print(f"‼️ {page} - Maʼlumot yoʻq yoki xatolik yuz berdi")
            time.sleep(2)
            continue
        
        # Items olish
        items = data.get("data", {}).get("items", [])
        
        for item in items:
            cursor.execute('''
            INSERT OR IGNORE INTO student_list (
                id, meta_id, university_code, university_name, full_name, short_name,
                first_name, second_name, third_name, gender_code, gender_name,
                birth_date, student_id_number, image, avg_gpa, avg_grade, total_credit,
                country_code, country_name, province_code, province_name,
                current_province_code, current_province_name, district_code, district_name,
                current_district_code, current_district_name, terrain_code, terrain_name,
                current_terrain_code, current_terrain_name, citizenship_code, citizenship_name,
                student_status_code, student_status_name, curriculum_id,
                education_form_code, education_form_name, education_type_code, education_type_name,
                payment_form_code, payment_form_name, student_type_code, student_type_name,
                social_category_code, social_category_name, accommodation_code, accommodation_name,
                department_id, department_code, department_name,
                department_structure_type_code, department_structure_type_name,
                department_locality_type_code, department_locality_type_name,
                department_parent, department_active,
                specialty_id, specialty_code, specialty_name,
                group_id, group_name, education_lang_code, education_lang_name,
                level_code, level_name, semester_id, semester_code, semester_name,
                education_year_code, education_year_name, education_year_current,
                year_of_enter, roommate_count, is_graduate, total_acload, other,
                created_at, updated_at, hash, validate_url
            ) VALUES (?,?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                item.get("id"),
                item.get("meta_id"),
                item.get("university", {}).get("code"),
                item.get("university", {}).get("name"),
                item.get("full_name"),
                item.get("short_name"),
                item.get("first_name"),
                item.get("second_name"),
                item.get("third_name"),
                item.get("gender", {}).get("code"),
                item.get("gender", {}).get("name"),
                item.get("birth_date"),
                item.get("student_id_number"),
                item.get("image"),
                item.get("avg_gpa"),
                item.get("avg_grade"),
                item.get("total_credit"),
                item.get("country", {}).get("code"),
                item.get("country", {}).get("name"),
                item.get("province", {}).get("code"),
                item.get("province", {}).get("name"),
                item.get("currentProvince", {}).get("name") if item.get("currentProvince") is not None else None,
                item.get("currentProvince", {}).get("name") if item.get("currentProvince") is not None else None,
                item.get("district", {}).get("code"),
                item.get("district", {}).get("name"),
                item.get("currentDistrict", {}).get("code") if item.get("currentDistrict") is not None else None,
                item.get("currentDistrict", {}).get("name") if item.get("currentDistrict") is not None else None,
                item.get("terrain", {}).get("code") if item.get("terrain") is not None else None,
                item.get("terrain", {}).get("name") if item.get("terrain") is not None else None,

                item.get("currentTerrain", {}).get("code") if item.get("currentTerrain") is not None else None,
                item.get("currentTerrain", {}).get("name") if item.get("currentTerrain") is not None else None,
                item.get("citizenship", {}).get("code"),
                item.get("citizenship", {}).get("name"),
                item.get("studentStatus", {}).get("code"),
                item.get("studentStatus", {}).get("name"),
                item.get("_curriculum"),
                item.get("educationForm", {}).get("code") if item.get("educationForm") is not None else None,
                item.get("educationForm", {}).get("name") if item.get("educationForm") is not None else None,
                

                item.get("educationType", {}).get("code"),
                item.get("educationType", {}).get("name"),
                item.get("paymentForm", {}).get("code"),
                item.get("paymentForm", {}).get("name"),
                item.get("studentType", {}).get("code"),
                item.get("studentType", {}).get("name"),
                item.get("socialCategory", {}).get("code"),
                item.get("socialCategory", {}).get("name"),
                item.get("accommodation", {}).get("code"),
                item.get("accommodation", {}).get("name"),
                item.get("department", {}).get("id"),
                item.get("department", {}).get("code"),
                item.get("department", {}).get("name"),
                item.get("department", {}).get("structureType", {}).get("code"),
                item.get("department", {}).get("structureType", {}).get("name"),
                item.get("department", {}).get("localityType", {}).get("code"),
                item.get("department", {}).get("localityType", {}).get("name"),
                item.get("department", {}).get("parent"),
                item.get("department", {}).get("active"),
                item.get("specialty", {}).get("id"),
                item.get("specialty", {}).get("code"),
                item.get("specialty", {}).get("name"),
                item.get("group", {}).get("id") if item.get("group") is not None else None,
                item.get("group", {}).get("name") if item.get("group") is not None else None,
         
                # item.get("group", {}).get("id"),
                # item.get("group", {}).get("name"),
                item.get("group", {}).get("educationLang", {}).get("code") if item.get("code") is not None else None,
                item.get("group", {}).get("educationLang", {}).get("name") if item.get("name") is not None else None,
                item.get("level", {}).get("code") if item.get("code") is not None else None,
                item.get("level", {}).get("name") if item.get("name") is not None else None,
                item.get("semester", {}).get("id"),
                item.get("semester", {}).get("code"),
                item.get("semester", {}).get("name"),
                item.get("educationYear", {}).get("code") if item.get("code") is not None else None,
                item.get("educationYear", {}).get("name") if item.get("name") is not None else None,
                item.get("educationYear", {}).get("current") if item.get("current") is not None else None,
                item.get("year_of_enter"),
                item.get("roommate_count"),
                item.get("is_graduate"),
                item.get("total_acload"),
                item.get("other"),
                item.get("created_at"),
                item.get("updated_at"),
                item.get("hash"),
                item.get("validateUrl")
            ))
        
        conn.commit()
        print(f"✅ {page} - sahifa yuklandi.")
        page += 1

    print(f"✅ Barcha {page_count} sahifa yuklandi.")

fetch_and_store_students()

# Ulanishni yopish
conn.close()
print("Barcha ma'lumotlar bazaga yozildi.")