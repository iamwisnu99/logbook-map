import os
import time
import random
import zipfile
import sys
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import glob
from PyPDF2 import PdfReader, PdfWriter
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardMarkup, InlineKeyboardButton, Bot
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from dotenv import load_dotenv
import asyncio
import nest_asyncio

# Load environment variables
load_dotenv()

# ===== KONFIGURASI TELEGRAM BOT =====
TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
TELEGRAM_CHAT_ID = os.getenv('TELEGRAM_CHAT_ID')

# ===== KONFIGURASI PATH LOGBOOK FINAL =====
BASE_LOGBOOK_PATH = r"D:\Documents\LOGBOOK MAP"


# ===== VARIABEL GLOBAL UNTUK KONTROL PROGRAM =====
program_state = {
    'running': False,
    'should_stop': False,
    'force_quit': False,
    'user_input': {
        'bulans_pending': [],
        'job_list': [],
        'kirim_telegram': False
    },
    'waiting_for_input': False,
    'current_step': None,
    'current_bulan_config': None,
    'message_id': None,
    'driver': None,
    'stop_after_logout': False

}

# ===== KONFIGURASI BULAN =====
BULAN_MAP = {
    "januari": 0, "februari": 1, "maret": 2, "april": 3,
    "mei": 4, "juni": 5, "juli": 6, "agustus": 7,
    "september": 8, "oktober": 9, "november": 10, "desember": 11
}

BULAN_NAMA = {
    0: "Januari", 1: "Februari", 2: "Maret", 3: "April",
    4: "Mei", 5: "Juni", 6: "Juli", 7: "Agustus",
    8: "September", 9: "Oktober", 10: "November", 11: "Desember"
}

BULAN_ABBR = {
    0: "Jan", 1: "Feb", 2: "Mar", 3: "Apr", 4: "Mei", 5: "Jun",
    6: "Jul", 7: "Ags", 8: "Sep", 9: "Okt", 10: "Nov", 11: "Des"
}

# ===== FUNGSI UNTUK MEMBACA DATA AKUN =====
def baca_data_akun(file_path="akun.xlsx"):
    """Membaca data akun dari file Excel"""
    try:
        df = pd.read_excel(file_path, dtype={'PIN': str, 'MID/Password': str})
        return df
    except FileNotFoundError:
        raise Exception(f"File '{file_path}' tidak ditemukan!")
    except Exception as e:
        raise Exception(f"Error membaca file: {str(e)}")

# ===== FUNGSI SETUP BROWSER =====
def setup_browser(download_path):
    """Setup browser dengan konfigurasi download otomatis"""
    chrome_options = Options()
    
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "plugins.plugins_disabled": ["Chrome PDF Viewer"]
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    chrome_options.add_argument(f"user-agent={user_agent}")

    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-geolocation")
    chrome_options.add_argument("--disable-media-stream")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    
    return driver

# ===== FUNGSI UNTUK MENUNGGU DOWNLOAD SELESAI =====
def tunggu_download_selesai(download_path, timeout=300):
    """Menunggu sampai file selesai diunduh"""
    pdf_files_sebelum = set(glob.glob(os.path.join(download_path, "*.pdf")))
    waktu_mulai = time.time()
    
    while time.time() - waktu_mulai < timeout:
        downloading = glob.glob(os.path.join(download_path, "*.crdownload"))
        
        if not downloading:
            pdf_files_sekarang = set(glob.glob(os.path.join(download_path, "*.pdf")))
            pdf_files_baru = pdf_files_sekarang - pdf_files_sebelum
            
            if pdf_files_baru:
                latest_file = max(pdf_files_baru, key=os.path.getmtime)
                return latest_file
        
        time.sleep(1)
    
    return None

# ===== FUNGSI UNTUK MENGHAPUS PASSWORD PDF =====
def hapus_password_pdf(file_path, password):
    """Menghapus password dari file PDF"""
    print(f"   → Memproses penghapusan password PDF...")
    
    try:
        reader = PdfReader(file_path)
        if reader.is_encrypted:
            if reader.decrypt(password):
                writer = PdfWriter()
                for page_num in range(len(reader.pages)):
                    writer.add_page(reader.pages[page_num])
                with open(file_path, 'wb') as output_file:
                    writer.write(output_file)
                print(f"   ✓ Password berhasil dihapus dari PDF")
                return True
            else:
                print(f"   ✗ Password salah, tidak dapat decrypt PDF")
                return False
        else:
            print(f"   ℹ PDF tidak memiliki password")
            return True
    except Exception as e:
        print(f"   ✗ Error saat menghapus password PDF: {str(e)}")
        return False

# ===== FUNGSI UNTUK RENAME FILE =====
def rename_file(file_path, nama, bulan_nama, tahun, output_folder):
    """Mengubah nama file dan memindahkannya ke folder output {Tahun}/{Bulan}"""
    nama_baru = f"{nama} - {bulan_nama} {tahun}.pdf"
    path_baru = os.path.join(output_folder, nama_baru)
    
    try:
        if os.path.exists(path_baru):
            os.remove(path_baru)
        shutil.move(file_path, path_baru)
        return path_baru
    except Exception as e:
        print(f"Gagal memindahkan/rename file: {e}")
        return file_path

# ===== FUNGSI PROSES LOGIN =====
def proses_login(driver, email, pin):
    """Melakukan proses login"""
    try:
        input_email = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Masukkan Nomor Ponsel atau Email']"))
        )
        input_email.clear()
        for character in email:
            input_email.send_keys(character)
            time.sleep(random.uniform(0.05, 0.15))
        time.sleep(0.5)
        
        input_pin = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Masukkan nomor PIN Anda']"))
        )
        input_pin.clear()
        for character in pin:
            input_pin.send_keys(character)
            time.sleep(random.uniform(0.05, 0.20))
        time.sleep(1)
        
        tombol_masuk = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'MASUK')]"))
        )
        tombol_masuk.click()
        time.sleep(4)
        return True
    except Exception as e:
        return False

# ===== FUNGSI NAVIGASI KE LOGBOOK =====
def navigasi_ke_logbook(driver, bulan_index, tahun_target):
    """Navigasi ke logbook dengan filter tahun dan klik ganda pada bulan untuk range"""
    try:
        wait = WebDriverWait(driver, 15)
        tahun_sekarang = str(datetime.now().year)
        perlu_filter = tahun_target.lower() != "saat ini" and tahun_target != tahun_sekarang

        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='Laporan Penjualan']"))).click()
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[text()='Logbook Bulanan']"))).click()
        time.sleep(4)

        if perlu_filter:
            print(f"   → Memilih rentang waktu: {BULAN_ABBR[bulan_index]} {tahun_target}")

            wait.until(EC.element_to_be_clickable((By.ID, "filter-rangemonth-picker"))).click()
            time.sleep(1)

            btn_prev = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-previous='true']")))
            btn_prev.click()
            time.sleep(1)

            nama_abbr = BULAN_ABBR[bulan_index]
            xpath_bulan = f"//button[contains(@class, 'mantine-MonthPickerInput-pickerControl') and text()='{nama_abbr}']"
            btn_bulan = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_bulan)))

            btn_bulan.click()
            print(f"     - Klik pertama pada {nama_abbr} (Start Range)")
            time.sleep(0.5)

            btn_bulan.click()
            print(f"     - Klik kedua pada {nama_abbr} (End Range)")
            
            time.sleep(4)

        target_id = "btnDownloadLb0" if perlu_filter else f"btnDownloadLb{bulan_index}"
        
        print(f"   → Menekan tombol unduh: {target_id}")
        tombol_download = wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[@data-testid='{target_id}']")))
        tombol_download.click()
        time.sleep(1)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@data-testid='btnFetchDownload']"))).click()
        return True
    except Exception as e:
        print(f"❌ Error Navigasi: {str(e)}")
        return False

# ===== FUNGSI PROSES LOGOUT =====
def proses_logout(driver):
    """Melakukan proses logout"""
    try:
        tombol_back1 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@data-testid='btnBack']"))
        )
        tombol_back1.click()
        time.sleep(2)
        tombol_back2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@data-testid='btnBack']"))
        )
        tombol_back2.click()
        time.sleep(2)
        tombol_logout = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@data-testid='btnLogout']"))
        )
        tombol_logout.click()
        time.sleep(1)
        tombol_keluar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'KELUAR')]"))
        )
        tombol_keluar.click()
        time.sleep(1)
        return True
    except Exception as e:
        return False

# ===== FUNGSI UNTUK ZIP FILE =====
def zip_folders(folder_list, zip_filename_base):
    """Membuat ZIP file dari daftar FOLDER"""
    tahun_sekarang = datetime.now().year
    zip_filename = f"{zip_filename_base} {tahun_sekarang}.zip"
    base_download_path = os.path.join(os.getcwd(), "downloads")
    zip_path = os.path.join(base_download_path, zip_filename)
    
    try:
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zipf:
            for folder_path in folder_list:
                folder_name = os.path.basename(folder_path)
                for root, _, files in os.walk(folder_path):
                    for file in files:
                        file_path_full = os.path.join(root, file)
                        arcname = os.path.join(folder_name, os.path.relpath(file_path_full, folder_path))
                        zipf.write(file_path_full, arcname)
        return zip_path
    except Exception as e:
        print(f"Gagal membuat ZIP: {e}")
        return None

# ===== FUNGSI PILIH AKUN =====
def pilih_akun_by_input(df_akun, pilihan):
    """Memilih akun berdasarkan input string"""
    pilihan = pilihan.strip().lower()
    if pilihan in ["all", "semua"]:
        return list(df_akun.index)
    elif pilihan.startswith("#"):
        try:
            exclude_str = pilihan.replace("#", "").replace(" ", "")
            exclude_numbers = [int(x) - 1 for x in exclude_str.split(",") if x]
            valid_exclude = [n for n in exclude_numbers if 0 <= n < len(df_akun)]
            selected_indices = [i for i in df_akun.index if i not in valid_exclude]
            return selected_indices
        except: return None
    else:
        try:
            selected_str = pilihan.replace(" ", "")
            selected_numbers = [int(x) - 1 for x in selected_str.split(",") if x]
            valid_numbers = [n for n in selected_numbers if 0 <= n < len(df_akun)]
            return valid_numbers if valid_numbers else None
        except: return None

# ===== TELEGRAM BOT HANDLERS =====
async def send_startup_message():
    """Kirim logo dan caption saat program pertama kali jalan"""
    print("Mengirim pesan startup ke Telegram...")
    try:
        bot = Bot(token=TELEGRAM_BOT_TOKEN)
        keyboard = [['Start Program', 'Stop'], ['Clean Up', 'Brute All']]
        reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        caption = (
            "Program telah aktif ✅\n\n"
            "Bot siap menerima perintah.\n"
            "Silakan ketik /start untuk melanjutkan."
        )
        with open('logo.png', 'rb') as photo:
            await bot.send_photo(
                chat_id=TELEGRAM_CHAT_ID,
                photo=photo,
                caption=caption,
                reply_markup=reply_markup
            )
        print("Pesan startup berhasil dikirim.")
    except FileNotFoundError:
        print("File 'logo.png' tidak ditemukan. Mengirim pesan teks saja.")
        await bot.send_message(
            chat_id=TELEGRAM_CHAT_ID,
            text=caption,
            reply_markup=reply_markup
        )
    except Exception as e:
        print(f"❌ Gagal mengirim pesan startup: {e}")
        print("Pastikan TELEGRAM_BOT_TOKEN dan TELEGRAM_CHAT_ID di .env sudah benar.")

async def clean_up_downloads(context: ContextTypes.DEFAULT_TYPE):
    """
    Menghapus semua file dan folder di dalam direktori 'downloads'
    dan mengirim laporan ke Telegram.
    """
    chat_id = TELEGRAM_CHAT_ID
    
    try:
        msg = await context.bot.send_message(chat_id=chat_id, text="🗑️ Memulai pembersihan folder 'downloads'...")
        
        base_download_path = os.path.join(os.getcwd(), "downloads")
        if not os.path.exists(base_download_path):
            await context.bot.edit_message_text(chat_id=chat_id, message_id=msg.message_id, text="ℹ️ Folder 'downloads' tidak ditemukan.")
            return

        items = glob.glob(os.path.join(base_download_path, "*"))
        
        if not items:
            await context.bot.edit_message_text(chat_id=chat_id, message_id=msg.message_id, text="ℹ️ Folder 'downloads' sudah bersih.")
            return

        deleted_files = 0
        deleted_folders = 0

        for item_path in items:
            try:
                if os.path.isfile(item_path):
                    os.remove(item_path)
                    deleted_files += 1
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                    deleted_folders += 1
            except Exception as e:
                print(f"Gagal menghapus {item_path}: {e}")
                await context.bot.send_message(chat_id=chat_id, text=f"⚠️ Gagal menghapus item: {os.path.basename(item_path)}")

        await context.bot.edit_message_text(
            chat_id=chat_id, 
            message_id=msg.message_id, 
            text=f"✅ Pembersihan selesai!\n\n- {deleted_files} file dihapus\n- {deleted_folders} folder dihapus"
        )
        
    except Exception as e:
        print(f"Error di clean_up_downloads: {e}")
        await context.bot.send_message(chat_id=chat_id, text=f"❌ Terjadi error saat proses pembersihan: {e}")

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /start"""
    chat_id = str(update.effective_chat.id)
    if chat_id != TELEGRAM_CHAT_ID:
        await update.message.reply_text("❌ Anda tidak memiliki akses ke bot ini.")
        return
    try:
        await context.bot.delete_message(chat_id=chat_id, message_id=update.message.message_id)
    except Exception: pass
    inline_keyboard = [[InlineKeyboardButton("Syarat dan Ketentuan", url="https://iamwisnu99.github.io/syarat-ketentuan/")]]
    inline_markup = InlineKeyboardMarkup(inline_keyboard)
    caption = (
        "❗ Sebelum melanjutkan penggunaan kode program ini, mohon luangkan waktu Anda untuk membaca dan memahami "
        "Syarat dan Ketentuan yang berlaku. Penggunaan kode ini menunjukkan persetujuan Anda terhadap semua ketentuan yang tertera ✍️\n\n"
        "Silahkan ketik 'Start Program' untuk memulai Program."
    )
    await update.message.reply_text(text=caption, reply_markup=inline_markup)
    program_state['message_id'] = None

# Fungsi helper baru untuk menanyakan akun per bulan
async def ask_for_accounts(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """
    Fungsi helper yang dipanggil untuk menanyakan konfigurasi akun
    untuk bulan berikutnya dalam antrian.
    """
    chat_id = str(update.effective_chat.id)
    message_id = program_state['message_id']

    if program_state['user_input']['bulans_pending']:
        bulan_info = program_state['user_input']['bulans_pending'].pop(0)
        program_state['current_bulan_config'] = bulan_info
        program_state['current_step'] = 'pilih_akun_per_bulan'
        
        try:
            df_akun = baca_data_akun()
            daftar_akun = "\n".join([f"{i+1}. {row['Nama']} ({row['Email']})" for i, row in df_akun.iterrows()])
            daftar_bulan_terpilih = ", ".join([b['bulan_info']['nama'] for b in program_state['user_input']['job_list']])

            konfirmasi_bulan_sebelumnya = ""
            if daftar_bulan_terpilih:
                konfirmasi_bulan_sebelumnya = f"✅ Bulan dikonfigurasi: *{daftar_bulan_terpilih}*\n\n"

            await context.bot.edit_message_text(
                chat_id=chat_id,
                message_id=message_id,
                text=(
                    f"{konfirmasi_bulan_sebelumnya}"
                    f"📋 *KONFIGURASI AKUN UNTUK: {bulan_info['nama'].upper()}*\n\n"
                    f"*DAFTAR AKUN YANG TERSEDIA:*\n{daftar_akun}\n\n"
                    f"*CARA PEMILIHAN:*\n"
                    f"- Pilih beberapa: 1,2,3\n"
                    f"- Kecuali tertentu: #1,#2,#3\n"
                    f"- Semua akun: all atau semua\n\n"
                    f"Akun mana yang ingin kamu unduh untuk bulan *{bulan_info['nama']}*?"
                ),
                parse_mode='Markdown'
            )
        except Exception as e:
            await context.bot.edit_message_text(chat_id=chat_id, message_id=message_id, text=f"❌ Error internal saat menyiapkan pertanyaan: {str(e)}")
            program_state['running'] = False
            program_state['waiting_for_input'] = False
            
    else:
        program_state['current_step'] = 'kirim_telegram'
        program_state['current_bulan_config'] = None
        
        await context.bot.edit_message_text(
            chat_id=chat_id,
            message_id=message_id,
            text=(
                "✅ Konfigurasi akun untuk semua bulan selesai.\n\n"
                "📋 *PERTANYAAN AKHIR*\n\n"
                "Apakah kamu ingin file yang diunduh dikirim via Telegram?\n\n"
                "Jawab: Ya / Tidak"
            ),
            parse_mode='Markdown'
        )

async def graceful_shutdown(context, message_id, reason):
    if program_state.get('driver'):
        try:
            program_state['driver'].quit()
        except:
            pass

    await context.bot.edit_message_text(
        chat_id=TELEGRAM_CHAT_ID,
        message_id=message_id,
        text=reason,
        parse_mode='Markdown'
    )

    program_state['running'] = False
    program_state['should_stop'] = False
    program_state['stop_after_logout'] = False
    program_state['force_quit'] = False
    program_state['waiting_for_input'] = False
    program_state['driver'] = None


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk semua pesan"""
    chat_id = str(update.effective_chat.id)
    if chat_id != TELEGRAM_CHAT_ID: return
    
    text = update.message.text.strip()
    try:
        await context.bot.delete_message(chat_id=chat_id, message_id=update.message.message_id)
    except Exception as e:
        print(f"Tidak bisa menghapus pesan user: {e}")

    message_id = program_state.get('message_id')
    
    if (not message_id and text.lower() not in ["stop", "brute all", "start program", "clean up"]):
        await update.message.reply_text("Silakan ketik /start dan 'Start Program' untuk memulai.", disable_notification=True)
        return

    if program_state['waiting_for_input'] and not message_id:
        await update.message.reply_text("Terjadi error. Silakan ketik 'Start Program' untuk mengulang.")
        program_state['waiting_for_input'] = False
        program_state['running'] = False
        return

    # ===== Handle Stop =====
    if text.lower() == "stop":
        if program_state['running']:
            program_state['should_stop'] = True
            program_state['stop_after_logout'] = True
            if message_id:
                await context.bot.edit_message_text(chat_id=chat_id, message_id=message_id, text="⏸️ Program akan dihentikan setelah akun saat ini selesai logout...")
        else:
            await update.message.reply_text("ℹ️ Program tidak sedang berjalan.")
        return

    # ===== Handle Brute All =====
    if text.lower() == "brute all":
        if program_state['running']:
            warning_text = "Brute All tidak dapat dilakukan, sebaiknya kirim perintah 'Stop' terlebih dahulu agar program dihentikan dengan aman."
            try:
                temp_msg = await update.message.reply_text(warning_text)
                await asyncio.sleep(5)
                await context.bot.delete_message(chat_id=chat_id, message_id=temp_msg.message_id)
            except Exception as e:
                print(f"Error saat mengirim/menghapus pesan Brute All: {e}")
        else:
            await update.message.reply_text("🛑 ATAS PERMINTAAN ANDA, Bot akan dimatikan (Brute All).\n\nHarap restart dari terminal dengan `python main.py`.")
            await asyncio.sleep(1)
            sys.exit("Brute All dipanggil saat program tidak berjalan.")
        return

    # ===== Handle Clean Up =====
    if text.lower() == "clean up":
        if program_state['running']:
            warning_text = "Clean Up tidak dapat dilakukan saat program sedang berjalan. Silakan kirim perintah 'Stop' terlebih dahulu."
            try:
                temp_msg = await update.message.reply_text(warning_text)
                await asyncio.sleep(5)
                await context.bot.delete_message(chat_id=chat_id, message_id=temp_msg.message_id)
            except Exception as e:
                print(f"Error saat mengirim/menghapus pesan Clean Up: {e}")
        else:
            await clean_up_downloads(context)
        return

    # ===== Handle Start Program =====
    if text.lower() == "start program":
        if program_state['running']:
            if message_id:
                await context.bot.edit_message_text(chat_id=chat_id, message_id=message_id, text="⚠️ Program sedang berjalan! Gunakan 'Stop' untuk menghentikan.")
            return

        program_state['running'] = True
        program_state['should_stop'] = False
        program_state['force_quit'] = False
        program_state['user_input'] = {'bulans_pending': [], 'job_list': [], 'kirim_telegram': False}
        program_state['waiting_for_input'] = True
        program_state['current_step'] = 'pilih_bulan'
        program_state['current_bulan_config'] = None
        
        message_text = (
            "📋 *PERTANYAAN 1: BULAN*\n\n"
            "Bulan apa yang ingin kamu unduh?\n\n"
            "Pilihan: Januari, Februari, Maret, April, Mei, Juni, Juli, Agustus, September, Oktober, November, Desember\n\n"
            "*(Bisa pilih lebih dari satu, pisahkan dengan koma, misal: Januari, Maret)*"
        )
        message = await context.bot.send_message(chat_id=chat_id, text=message_text, parse_mode='Markdown')
        program_state['message_id'] = message.message_id
        return

    # ===== Handle Input Konfigurasi =====
    if program_state['waiting_for_input'] and message_id:

        if program_state['current_step'] == 'pilih_bulan':
            bulan_inputs = [b.strip().lower() for b in text.split(',')]
            selected_bulans = []
            invalid_bulans = []
            
            for b in bulan_inputs:
                if b in BULAN_MAP:
                    bulan_data = {'index': BULAN_MAP[b], 'nama': BULAN_NAMA[BULAN_MAP[b]]}
                    if bulan_data not in selected_bulans:
                        selected_bulans.append(bulan_data)
                else:
                    invalid_bulans.append(b)

            if invalid_bulans:
                await context.bot.edit_message_text(
                    chat_id=chat_id, message_id=message_id,
                    text=(
                        f"❌ Bulan tidak valid: *{', '.join(invalid_bulans)}*!\n\n"
                        "📋 *PERTANYAAN 1: BULAN*\n\n"
                        "Bulan apa yang ingin kamu unduh?\n"
                        "*(Pisahkan dengan koma jika lebih dari satu)*"
                    ), parse_mode='Markdown'
                )
                return
            
            if not selected_bulans:
                await context.bot.edit_message_text(
                    chat_id=chat_id, message_id=message_id,
                    text=(
                        "❌ Anda harus memasukkan setidaknya satu nama bulan yang valid.\n\n"
                        "📋 *PERTANYAAN 1: BULAN*\n\n"
                        "Bulan apa yang ingin kamu unduh?"
                    ), parse_mode='Markdown'
                )
                return

            program_state['user_input']['bulans_pending'] = selected_bulans
            program_state['user_input']['job_list'] = []
            program_state['current_step'] = 'pilih_tahun'
            await context.bot.edit_message_text(
                chat_id=chat_id, message_id=message_id,
                text="📋 *PERTANYAAN 2: TAHUN*\n\nTahun berapa yang ingin kamu pilih?\n*(Contoh: 2025 atau Saat Ini)*",
                parse_mode='Markdown'
            )
            return

        elif program_state['current_step'] == 'pilih_tahun':
            input_user = text.strip()

            if input_user.lower() == "saat ini" or (input_user.isdigit() and len(input_user) == 4):
                program_state['user_input']['tahun_terpilih'] = input_user
                await ask_for_accounts(update, context)
            else:
                await context.bot.edit_message_text(
                    chat_id=chat_id, message_id=message_id,
                    text="❌ Input tidak valid!\nKetik **Saat Ini** atau angka tahun (Contoh: **2024**).",
                    parse_mode='Markdown'
                )
            return

        elif program_state['current_step'] == 'pilih_akun_per_bulan':
            try:
                df_akun = baca_data_akun()
                selected_indices = pilih_akun_by_input(df_akun, text)
                
                if selected_indices is None:
                    await context.bot.edit_message_text(
                        chat_id=chat_id, message_id=message_id,
                        text=(
                            f"❌ Format tidak valid! Coba lagi untuk bulan *{program_state['current_bulan_config']['nama']}*.\n"
                            f"Gunakan format: 1,2,3 atau all atau #1,#2"
                        ), parse_mode='Markdown'
                    )
                    await asyncio.sleep(2)
                    await ask_for_accounts(update, context)
                    return

                current_job = {
                    'bulan_info': program_state['current_bulan_config'], 
                    'selected_indices': selected_indices,
                    'df_akun': df_akun
                }
                program_state['user_input']['job_list'].append(current_job)

                await ask_for_accounts(update, context)
                return
                
            except Exception as e:
                await context.bot.edit_message_text(chat_id=chat_id, message_id=message_id, text=f"❌ Error: {str(e)}")
                program_state['running'] = False
                program_state['waiting_for_input'] = False
            return

        elif program_state['current_step'] == 'kirim_telegram':
            if text.lower() in ['ya', 'tidak']:
                program_state['user_input']['kirim_telegram'] = (text.lower() == 'ya')
                program_state['waiting_for_input'] = False

                confirmation_text = "✅ *KONFIGURASI SELESAI*\n\nProgram akan menjalankan tugas berikut:\n\n"
                
                for job in program_state['user_input']['job_list']:
                    bulan_nama = job['bulan_info']['nama']
                    indices = job['selected_indices']
                    df_akun = job['df_akun']
                    
                    df_terpilih = df_akun.iloc[indices]
                    nama_akun_terpilih = ", ".join(df_terpilih['Nama'].tolist())
                    if len(indices) == len(df_akun):
                        nama_akun_terpilih = f"Semua Akun ({len(indices)})"
                    
                    confirmation_text += f"🗓️ *Bulan: {bulan_nama}*\n   👥 Akun: {nama_akun_terpilih}\n\n"
                
                confirmation_text += f"📤 Kirim Telegram: {'Ya' if program_state['user_input']['kirim_telegram'] else 'Tidak'}\n\n"
                confirmation_text += "⏳ Program akan segera dimulai..."
                
                await context.bot.edit_message_text(
                    chat_id=chat_id,
                    message_id=message_id,
                    text=confirmation_text,
                    parse_mode='Markdown'
                )

                await run_main_process(update, context)
                
            else:
                await context.bot.edit_message_text(
                    chat_id=chat_id, message_id=message_id,
                    text=(
                        "❌ Jawaban tidak valid! Silakan jawab 'Ya' atau 'Tidak'\n\n"
                        "📋 *PERTANYAAN AKHIR*\n\n"
                        "Apakah kamu ingin file yang diunduh dikirim via Telegram?"
                    ), parse_mode='Markdown'
                )
            return

async def run_main_process(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Menjalankan proses utama pengunduhan"""
    try:
        job_list = program_state['user_input']['job_list']
        tahun_user = program_state['user_input'].get('tahun_terpilih', 'Saat Ini')
        if tahun_user.lower() == "saat ini":
            tahun_user = str(datetime.now().year)
        kirim_telegram = program_state['user_input']['kirim_telegram']
        df_akun = baca_data_akun()

        message_id = program_state['message_id']
        if not message_id:
            msg = await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="Memulai... (Error: message_id hilang, membuat pesan baru)")
            program_state['message_id'] = msg.message_id
            message_id = msg.message_id

        base_download_path = os.path.join(os.getcwd(), "downloads_temp")
        os.makedirs(base_download_path, exist_ok=True)


        driver = setup_browser(base_download_path)
        program_state['driver'] = driver
        
        folders_created_for_zip = []
        total_files_downloaded = 0

        for job in job_list:
            if program_state['should_stop'] or program_state['force_quit']:
                break

            bulan_info = job['bulan_info']
            job_selected_indices = job['selected_indices']
            
            bulan_index = bulan_info['index']
            bulan_nama = bulan_info['nama']
            tahun_target = program_state['user_input'].get('tahun_terpilih', datetime.now().year)

            folder_bulan_tahun = f"{bulan_nama} {tahun_user}"

            output_folder_path = os.path.join(
                BASE_LOGBOOK_PATH,
                str(tahun_user),
                folder_bulan_tahun
            )

            os.makedirs(output_folder_path, exist_ok=True)

            
            if output_folder_path not in folders_created_for_zip:
                folders_created_for_zip.append(output_folder_path)

            df_akun_terpilih = df_akun.iloc[job_selected_indices].reset_index(drop=True)
            total_akun = len(df_akun_terpilih)

            await context.bot.edit_message_text(
                chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                text=(
                    f"🔄 *Program Sedang Berjalan* ✅\n\n"
                    f"🗓️ *Memproses Bulan: {bulan_nama}*\n"
                    f"📧 Email: Memulai...\n"
                    f"📊 Progress: 0/8\n"
                    f"👤 Akun: 0/{total_akun}"
                ), parse_mode='Markdown'
            )

            for index, row in df_akun_terpilih.iterrows():
                if program_state['should_stop'] or program_state['force_quit']:
                    break
                
                email = row['Email']
                pin = str(row['PIN']).strip()
                nama = row['Nama']
                password_pdf = str(row['MID/Password']).strip() if 'MID/Password' in row and pd.notna(row['MID/Password']) else ""
                
                try:
                    await context.bot.edit_message_text(
                        chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                        text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 1/8 - Membuka halaman login\n👤 Akun: {index+1}/{total_akun}",
                        parse_mode='Markdown'
                    )
                    driver.get("https://subsiditepatlpg.mypertamina.id/merchant-login")
                    time.sleep(5)
                    
                    await context.bot.edit_message_text(
                        chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                        text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 2/8 - Melakukan login\n👤 Akun: {index+1}/{total_akun}",
                        parse_mode='Markdown'
                    )
                    if not proses_login(driver, email, pin):
                        await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"❌ Gagal login untuk {nama} ({email}). Melanjutkan...")
                        continue
                    
                    await context.bot.edit_message_text(
                        chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                        text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 4/8 - Navigasi ke logbook\n👤 Akun: {index+1}/{total_akun}",
                        parse_mode='Markdown'
                    )
                    if not navigasi_ke_logbook(driver, bulan_index, tahun_user):
                        await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"❌ Gagal navigasi untuk {nama}. Logout...")
                        proses_logout(driver)
                        continue
                    
                    await context.bot.edit_message_text(
                        chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                        text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 5/8 - Mengunduh file\n👤 Akun: {index+1}/{total_akun}",
                        parse_mode='Markdown'
                    )
                    file_downloaded = tunggu_download_selesai(base_download_path)
                    
                    if file_downloaded:
                        await context.bot.edit_message_text(
                            chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                            text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 6/8 - Rename & Pindah file\n👤 Akun: {index+1}/{total_akun}",
                            parse_mode='Markdown'
                        )
                        file_baru = rename_file(file_downloaded, nama, bulan_nama, tahun_user, output_folder_path)
                        
                        if password_pdf:
                            await context.bot.edit_message_text(
                                chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                                text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 7/8 - Menghapus password PDF\n👤 Akun: {index+1}/{total_akun}",
                                parse_mode='Markdown'
                            )
                            hapus_password_pdf(file_baru, password_pdf)
                        
                        total_files_downloaded += 1
                    else:
                        await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"❌ Gagal mengunduh file untuk {nama}.")
                    
                    await context.bot.edit_message_text(
                        chat_id=TELEGRAM_CHAT_ID, message_id=message_id,
                        text=f"🔄 *Bulan: {bulan_nama}*\n\n📧 Email: {email}\n📊 Progress: 8/8 - Logout\n👤 Akun: {index+1}/{total_akun}",
                        parse_mode='Markdown'
                    )
                    proses_logout(driver)

                    if program_state['stop_after_logout']:
                        await graceful_shutdown(
                            context,
                            message_id,
                            "⏸️ Program dihentikan oleh user setelah logout akun terakhir."
                        )
                        return


                except Exception as e:
                    await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"❌ Error pada akun {nama}: {str(e)}")
                    try:
                        proses_logout(driver)
                    except:
                        driver.quit()
                        driver = setup_browser(base_download_path)
                        program_state['driver'] = driver
                    continue
                
                if program_state['should_stop'] or program_state['force_quit']:
                    break
            
        driver.quit()
        program_state['driver'] = None

        final_text = ""
        if program_state['force_quit']:
            final_text = "🛑 *Program dihentikan secara paksa!*"
        elif program_state['should_stop']:
            final_text = f"⏸️ *Program dihentikan oleh user total file diunduh: {total_files_downloaded}"
        else:
            final_text = f"✅ *SEMUA PROSES SELESAI*\n\n📦 Total file berhasil diunduh: {total_files_downloaded}"
        
        await context.bot.edit_message_text(chat_id=TELEGRAM_CHAT_ID, message_id=message_id, text=final_text, parse_mode='Markdown')

        if kirim_telegram and folders_created_for_zip and not program_state['force_quit']:
            zip_status_msg = await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="📦 Membuat file ZIP (level kompresi 9)...")
            
            zip_base_name = "Logbook Multi-Bulan"
            if len(folders_created_for_zip) == 1:
                zip_base_name = os.path.basename(folders_created_for_zip[0])
            
            zip_path = zip_folders(folders_created_for_zip, zip_base_name)
            
            if zip_path:
                await context.bot.edit_message_text(chat_id=TELEGRAM_CHAT_ID, message_id=zip_status_msg.message_id, text="📤 Mengirim file ZIP...")
                
                try:
                    with open(zip_path, 'rb') as zip_file:
                        await context.bot.send_document(
                            chat_id=TELEGRAM_CHAT_ID,
                            document=zip_file,
                            filename=os.path.basename(zip_path),
                            caption=f"📦 *File Logbook*\n\n✅ Berisi {total_files_downloaded} file PDF dari {len(folders_created_for_zip)} bulan.",
                            pool_timeout=1200,
                            connect_timeout=60,
                            parse_mode='Markdown'
                        )
                    await context.bot.delete_message(chat_id=TELEGRAM_CHAT_ID, message_id=zip_status_msg.message_id)

                    cleanup_msg = await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text="✅ ZIP terkirim! Menjadwalkan pembersihan folder 'downloads' dalam 10 detik...")
                    await asyncio.sleep(10) 

                    await context.bot.edit_message_text(chat_id=TELEGRAM_CHAT_ID, message_id=cleanup_msg.message_id, text="... membersihkan file dan folder ...")
                    
                    deleted_folders = 0
                    deleted_files = 0
                    
                    for folder_path in folders_created_for_zip:
                        try:
                            shutil.rmtree(folder_path)
                            deleted_folders += 1
                        except Exception as e:
                            print(f"Gagal menghapus folder {folder_path}: {e}")
                    try:
                        os.remove(zip_path)
                        deleted_files += 1
                    except Exception as e:
                        print(f"Gagal menghapus file ZIP {zip_path}: {e}")

                    await context.bot.edit_message_text(chat_id=TELEGRAM_CHAT_ID, message_id=cleanup_msg.message_id, text=f"🗑️ Pembersihan selesai! {deleted_folders} folder dan {deleted_files} file ZIP telah dihapus.")

                except Exception as e:
                    pass
            else:
                await context.bot.edit_message_text(chat_id=TELEGRAM_CHAT_ID, message_id=zip_status_msg.message_id, text="❌ Gagal membuat file ZIP")

        else:
            await context.bot.send_message(
                chat_id=TELEGRAM_CHAT_ID,
                text=(
                    "💾 *File disimpan secara lokal*\n\n"
                    "📁 *Folder Penyimpanan:*\n"
                    f"`{output_folder_path}`"
                ),
                parse_mode='Markdown'
            )

        program_state['running'] = False
        program_state['should_stop'] = False
        program_state['force_quit'] = False
        program_state['waiting_for_input'] = False
        program_state['message_id'] = None
        
    except Exception as e:
        if program_state['message_id']:
            await context.bot.edit_message_text(chat_id=TELEGRAM_CHAT_ID, message_id=program_state['message_id'], text=f"❌ *Error Fatal:* {str(e)}", parse_mode='Markdown')
        else:
             await context.bot.send_message(chat_id=TELEGRAM_CHAT_ID, text=f"❌ *Error Fatal:* {str(e)}", parse_mode='Markdown')
        
        if program_state['driver']:
            try: driver.quit()
            except: pass
    except StopIteration:
        pass     
        program_state['running'] = False
        program_state['should_stop'] = False
        program_state['stop_after_logout'] = False
        program_state['force_quit'] = False
        program_state['message_id'] = None

# ===== FUNGSI UTAMA =====
def main():
    """Fungsi utama program"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("❌ Error: TELEGRAM_BOT_TOKEN dan TELEGRAM_CHAT_ID harus diisi di file .env")
        return
    
    nest_asyncio.apply()
    
    print("="*50); print("🤖 TELEGRAM BOT PENGUNDUHAN LOGBOOK"); print("="*50)

    try:
        asyncio.run(send_startup_message())
    except Exception as e:
        print(f"Gagal di asyncio.run(send_startup_message): {e}")

    print(f"\n✅ Bot Token: {TELEGRAM_BOT_TOKEN[:20]}..."); print(f"✅ Chat ID: {TELEGRAM_CHAT_ID}")
    print("\n🚀 Bot sedang berjalan..."); print("📱 Silakan buka Telegram dan ketik /start untuk memulai\n"); print("Tekan Ctrl+C untuk menghentikan bot\n"); print("="*50)
    
    application = Application.builder().token(TELEGRAM_BOT_TOKEN).build()
    
    application.add_handler(CommandHandler("start", start_command))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    try:
        application.run_polling(allowed_updates=Update.ALL_TYPES)
    except KeyboardInterrupt:
        print("\n\n🛑 Bot dihentikan oleh user")
    except Exception as e:
        print(f"\n\n❌ Error: {str(e)}")

if __name__ == "__main__":
    main()