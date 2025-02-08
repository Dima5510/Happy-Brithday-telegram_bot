import pandas as pd
from telegram import Bot
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import os
import schedule
import time

# Bot tokenini o'rnating
BOT_TOKEN = '7758803370:AAFd-x1xKxWMMb3X88HzHQcGc8noAbw323k'
CHANNEL_ID = '-1002407387026'

# Excel faylni o'qish
def load_employee_data():
    try:
        employees = pd.read_excel('employees.xlsx')
        print("Excel faylidan olingan ustunlar:", employees.columns)  # Ustun nomlarini chop etish
        employees['Tug‘ilgan kuni'] = pd.to_datetime(employees['Tug‘ilgan kuni'], errors='coerce').dt.strftime('%d.%m')
        return employees
    except Exception as e:
        print(f"Xatolik: {e}")
        return None

# Tabrik kartasini yaratish
def create_birthday_card(name, photo_path, output_path):
    try:
        background = Image.open("background.jpg").convert("RGBA")
        background = background.resize((1080, 1920))

        if os.path.exists(photo_path):
            user_photo = Image.open(photo_path).convert("RGBA")
            user_photo = user_photo.resize((500, 500))
        else:
            print(f"Rasm topilmadi: {photo_path}")
            return False

        mask = Image.new('L', (500, 500), 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, 500, 500), fill=255)
        user_photo.putalpha(mask)
        background.paste(user_photo, (290, 200), user_photo)

        draw = ImageDraw.Draw(background)
        font = ImageFont.truetype("arial.ttf", 80)
        text = f"🎉 Tug‘ilgan kuning bilan, {name}! 🎉"
        text_width, _ = draw.textsize(text, font=font)
        draw.text(((1080 - text_width) / 2, 1500), text, font=font, fill="white")

        background.save(output_path)
        return True
    except Exception as e:
        print(f"Tabrik kartasini yaratishda xatolik: {e}")
        return False

# Tug'ilgan kunni tekshirish va tabrik yuborish
def check_birthdays():
    employees = load_employee_data()
    if employees is None:
        return

    today = datetime.now().strftime('%d.%m')
    print(f"Bugungi sana: {today}")

    for _, row in employees.iterrows():
        birthday = row['Tug\'ilgan kuni']
        if pd.isna(birthday):
            continue

        if birthday == today:
            name = row['Ismi']
            department = row['Bo\'lim']
            photo_path = row['photo_path']
            card_path = f"{name}_birthday_card.png"

            print(f"Tug'ilgan kun: {name}, {birthday}")

            if not create_birthday_card(name, photo_path, card_path):
                print(f"Tabrik kartasini yaratishda xatolik: {name}")
                continue

            bot = Bot(token=BOT_TOKEN)
            try:
                with open(card_path, 'rb') as photo_file:
                    bot.send_photo(chat_id=CHANNEL_ID, photo=photo_file, caption=f"🎉 {name}, tug‘ilgan kuning bilan! 🎉\n\n{department} bo‘limining azizi!")
                print(f"Tabrik yuborildi: {name}")
            except Exception as e:
                print(f"Xabar yuborishda xatolik: {e}")

            os.remove(card_path)

# Scheduler: Har kuni belgilangan vaqtda ishga tushirish
def run_scheduler():
    schedule.every().day.at("14:58").do(check_birthdays)
    while True:
        schedule.run_pending()
        time.sleep(1)

if __name__ == "__main__":
    print("Bot ishga tushdi...")
    run_scheduler()