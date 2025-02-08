import pandas as pd
from telegram import Bot
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import os
import schedule
import time
import asyncio

# Bot token va kanal ID
BOT_TOKEN = '7830646110:AAHSyzIpz-Wo_mrTeBKLqIiJetKOsyJ9f_g'
CHANNEL_ID = '-1002407387026'

# Excel faylni o'qish
def load_employee_data():
    try:
        employees = pd.read_excel('employees.xlsx')
        employees.columns = employees.columns.str.strip()
        print("Excel ustun nomlari:", employees.columns)
        
        # Tug'ilgan kun sanalarini to'g'ri formatlash
        if "Tug'ilgan_kuni" in employees.columns:
            employees['Tug\'ilgan_kuni'] = pd.to_datetime(
                employees['Tug\'ilgan_kuni'], errors='coerce'
            ).dt.strftime('%d.%m')
        else:
            print("⚠️ Excel faylda 'Tug'ilgan_kuni' ustuni topilmadi!")
            return None
        
        return employees
    except Exception as e:
        print(f"❌ Xatolik: {e}")
        return None

# Tabrik kartasini yaratish
def create_birthday_card(name, photo_path, output_path):
    try:
        # Fon rasmini ochish
        background = Image.open("background.jpg").convert("RGBA")
        background = background.resize((1080, 1920))  # Kartani o'lchamini sozlash

        # Hodimning rasmini ochish va aylana shaklida kesish
        if os.path.exists(photo_path):
            user_photo = Image.open(photo_path).convert("RGBA")
            user_photo = user_photo.resize((700, 700))  # Rasm o'lchamini sozlash
        else:
            print(f"⚠️ Rasm topilmadi: {photo_path}")
            return False

        # Aylana shaklidagi mask yaratish
        mask = Image.new('L', (700, 700), 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, 700, 700), fill=255)
        user_photo.putalpha(mask)

        # Rasmni fon ustiga joylashtirish
        position_x = background.width - user_photo.width - 200  # O'ng tomondan 100 piksel masofa
        position_y = (background.height - user_photo.height) // 2  # Vertikal markazlash
        background.paste(user_photo, (position_x, position_y), user_photo)

        # Matnni qo'shish
        # draw = ImageDraw.Draw(background)
        # font = ImageFont.truetype("arial.ttf", 20)  # Shrift faylini tanlash
        # text = f"🎉 Tug'ilgan kuning bilan, {name}! 🎉"
        
        # # Pillow 10.0.0+ uchun textsize o'rniga textbbox ishlatish
        # text_bbox = draw.textbbox((0, 0), text, font=font)
        # text_width = text_bbox[2] - text_bbox[0]
        # text_height = text_bbox[3] - text_bbox[1]
        
        # draw.text(((1080 - text_width) / 2, 1500), text, font=font, fill="white")

        # Natijaviy kartani saqlash
        background.save(output_path)
        return True
    except Exception as e:
        print(f"❌ Tabrik kartasini yaratishda xatolik: {e}")
        return False

# Tug'ilgan kunlarni tekshirish va tabrik yuborish
async def check_birthdays():
    employees = load_employee_data()
    if employees is None:
        return
    
    today = datetime.now().strftime('%d.%m')
    print(f"🔍 Bugungi sana: {today}")
    
    for _, row in employees.iterrows():
        birthday = row['Tug\'ilgan_kuni']
        if pd.isna(birthday) or birthday != today:
            continue
        
        name = row["Ismi"]
        department = row['Bo\'limi']
        photo_path = row['Photo_path']
        card_path = f"{name}_birthday_card.png"
        
        print(f"🎂 Tug'ilgan kuni: {name}, {birthday}")
        
        if not create_birthday_card(name, photo_path, card_path):
            print(f"❌ Tabrik kartasini yaratishda xatolik: {name}")
            continue
        
        bot = Bot(token=BOT_TOKEN)
        try:
            with open(card_path, 'rb') as photo_file:
                await bot.send_photo(chat_id=CHANNEL_ID, photo=photo_file, caption=f"🎉 {name}, tug'ilgan kuniz bilan! 🎉\n\n{department} bo'limining azizi!")
            print(f"✅ Tabrik yuborildi: {name}")
        except Exception as e:
            print(f"❌ Xabar yuborishda xatolik: {e}")
        os.remove(card_path)

# Har kuni belgilangan vaqtda ishga tushirish
def run_scheduler():
    schedule.every().day.at("11:38").do(asyncio.run, check_birthdays())  # Test uchun yaqin vaqtga o'zgartiring
    while True:
        print(f"🕒 Hozirgi vaqt: {datetime.now().strftime('%H:%M:%S')}")
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    print("🚀 Bot ishga tushdi...")
    run_scheduler()