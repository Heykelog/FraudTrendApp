from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import os
import functools
import datetime
import xlsxwriter
from io import BytesIO
import random
from xlsxwriter.utility import xl_rowcol_to_cell
from datetime import timedelta
from flask_wtf.csrf import CSRFProtect
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import threading
import time

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'ürettiğiniz_güçlü_bir_anahtar')  # Güvenlik için önemli

EXCEL_DOSYASI = 'fraud_trendi.xlsx'
DEFAULT_SIFRE = 'TF2025'
ADMIN_SIFRE = 'admin123'  # Admin girişi için şifre
KULLANICI_DOSYASI = 'kullanicilar.txt'
EMAIL_AYARLARI_DOSYASI = 'email_ayarlari.json'

# Varsayılan e-posta ayarları
DEFAULT_EMAIL_AYARLARI = {
    "smtp_sunucu": "smtp.gmail.com",
    "smtp_port": 587,
    "kullanici_adi": "",
    "sifre": "",
    "gonderen_email": "",
    "alici_email": "",
    "bildirim_saati": 8,  # Varsayılan olarak sabah 8:00'de gönder
    "bildirim_aktif": False
}

# E-posta ayarları dosyasını kontrol et, yoksa oluştur
def email_ayarlarini_kontrol_et():
    if not os.path.exists(EMAIL_AYARLARI_DOSYASI):
        with open(EMAIL_AYARLARI_DOSYASI, 'w', encoding='utf-8') as f:
            json.dump(DEFAULT_EMAIL_AYARLARI, f, ensure_ascii=False, indent=4)
    
    try:
        with open(EMAIL_AYARLARI_DOSYASI, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"E-posta ayarları dosyası okuma hatası: {e}")
        return DEFAULT_EMAIL_AYARLARI

# E-posta ayarlarını güncelle
def email_ayarlarini_guncelle(yeni_ayarlar):
    try:
        # JSON dosyasına yaz
        with open(EMAIL_AYARLARI_DOSYASI, 'w', encoding='utf-8') as f:
            json.dump(yeni_ayarlar, f, ensure_ascii=False, indent=4)
        
        print(f"E-posta ayarları güncellendi: {yeni_ayarlar}")
        return True
    except Exception as e:
        print(f"E-posta ayarları güncelleme hatası: {e}")
        return False

# E-posta gönderme fonksiyonu
def email_gonder(konu, icerik, alici_email=None):
    ayarlar = email_ayarlarini_kontrol_et()
    
    if not ayarlar["bildirim_aktif"]:
        print("E-posta bildirimleri devre dışı.")
        return False
    
    try:
        # E-posta ayarlarını al
        smtp_sunucu = ayarlar["smtp_sunucu"]
        smtp_port = ayarlar["smtp_port"]
        kullanici_adi = ayarlar["kullanici_adi"]
        sifre = ayarlar["sifre"]
        gonderen_email = ayarlar["gonderen_email"]
        
        # Alıcı e-posta adresini belirle
        if alici_email is None:
            alici_email = ayarlar["alici_email"]
        
        # E-posta oluştur
        mesaj = MIMEMultipart()
        mesaj["From"] = gonderen_email
        mesaj["To"] = alici_email
        mesaj["Subject"] = konu
        
        # İçerik ekle
        mesaj.attach(MIMEText(icerik, "html"))
        
        # SMTP bağlantısı kur ve e-posta gönder
        with smtplib.SMTP(smtp_sunucu, smtp_port) as server:
            server.starttls()
            server.login(kullanici_adi, sifre)
            server.send_message(mesaj)
        
        print(f"E-posta başarıyla gönderildi: {alici_email}")
        return True
    
    except Exception as e:
        print(f"E-posta gönderme hatası: {e}")
        return False

# Vaka bildirim e-postası içeriğini hazırla
def vaka_bildirim_email_icerigi_olustur(vaka):
    html = """
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <div style="max-width: 800px; margin: 0 auto; padding: 20px;">
            <div style="background-color: #2c3e50; color: white; padding: 10px 20px;">
                <h1 style="color: white; margin: 0;">Yeni Vaka Bildirimi</h1>
            </div>
            <div style="padding: 20px; border: 1px solid #ddd;">
                <div style="margin-bottom: 20px;">
                    <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Vaka Özeti</h2>
                    <table style="width: 100%; border-collapse: collapse;">
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Vaka Türü</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Kanal</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Dönem</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Müşteri/Üye İşyeri</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Müşteri No/TCKN/Kart No</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Tespit Tarihi</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                    </table>
                </div>
                
                <div style="margin-bottom: 20px;">
                    <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Finansal Bilgiler</h2>
                    <table style="width: 100%; border-collapse: collapse;">
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Kullanılan Tutar</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{} TL</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Müşteri Zararı</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{} TL</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Banka Zararı</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{} TL</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Kurtarılan Tutar</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{} TL</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Toplam Risk</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{} TL</td>
                        </tr>
                    </table>
                </div>
                
                <div style="margin-bottom: 20px;">
                    <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Vaka Detayları</h2>
                    <p>{}</p>
                </div>
                
                <div style="margin-bottom: 20px;">
                    <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Analist Notu</h2>
                    <p>{}</p>
                </div>
                
                <div style="margin-bottom: 20px;">
                    <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">İşlem Bilgileri</h2>
                    <table style="width: 100%; border-collapse: collapse;">
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">İşlem Türü</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Vaka Giren Kişi</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                        <tr>
                            <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">İnceleyen Kişi</th>
                            <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
                        </tr>
                    </table>
                </div>
                
                {}
            </div>
        </div>
    </body>
    </html>
    """.format(
        vaka.get('Vaka Türü', '-'),
        vaka.get('Kanal', '-'),
        vaka.get('Çeyreklik Dönem', '-'),
        vaka.get('Ad-Soyad / Üye İş Yeri Adı', '-'),
        vaka.get('Müşteri No / TCKN / Kart No', '-'),
        vaka.get('Sahte İşlem / Risk Tespiti Tarihi', '-'),
        vaka.get('Kullanılan Tutar', '-'),
        vaka.get('Müşteri Zararı', '-'),
        vaka.get('Banka Zararı', '-'),
        vaka.get('Kurtarılan Tutar', '-'),
        vaka.get('Toplam Risk (TL)', '-'),
        vaka.get('Vaka Detayı', '-'),
        vaka.get('Analist Notu', '-'),
        vaka.get('İşlem Türü', '-'),
        vaka.get('Vaka Giren Kişi', '-'),
        vaka.get('İnceleyen Kişi', '-'),
        _kart_bilgisi_olustur(vaka) if vaka.get('Kanal') == 'Kart' else ''
    )
    return html

def _kart_bilgisi_olustur(vaka):
    """Kart bilgilerini içeren HTML bölümünü oluşturur"""
    return """
    <div style="margin-bottom: 20px;">
        <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Kart Bilgileri</h2>
        <table style="width: 100%; border-collapse: collapse;">
            <tr>
                <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">CPP Tespiti</th>
                <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
            </tr>
            <tr>
                <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Pos/ATM</th>
                <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
            </tr>
            <tr>
                <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Kopyalama Onus/Not Onus</th>
                <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
            </tr>
            <tr>
                <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">Kopyalama Yeri/Şüpheli İşyeri</th>
                <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
            </tr>
            <tr>
                <th style="background-color: #f5f5f5; text-align: left; padding: 10px; border: 1px solid #ddd;">ATM Kart Kopyalama Yöntemi</th>
                <td style="padding: 10px; border: 1px solid #ddd;">{}</td>
            </tr>
        </table>
    </div>
    """.format(
        vaka.get('CPP Tespiti', '-'),
        vaka.get('Pos/ATM', '-'),
        vaka.get('Kopyalama Onus/Not Onus', '-'),
        vaka.get('Kopyalama Yeri/ Şüpheli İş yeri', '-'),
        vaka.get('ATM Kart Kopyalama Yöntemi', '-')
    )

# Çeyreklik dönem seçeneklerini hazırla
def get_ceyreklik_donemler():
    """Dönem seçeneklerini yıl ve çeyrek olarak ayrı listeler halinde döndürür"""
    simdiki_yil = datetime.datetime.now().year
    
    # Yıl seçenekleri (mevcut yıl öncesindeki 2 yıl ve sonrasındaki 2 yıl)
    yillar = list(range(simdiki_yil - 2, simdiki_yil + 3))
    
    # Çeyrek seçenekleri
    ceyrekler = ["Q1", "Q2", "Q3", "Q4"]
    
    return {
        'yillar': yillar,
        'ceyrekler': ceyrekler,
        'simdiki_yil': simdiki_yil
    }

# Kullanıcı sınıfı
class Kullanici:
    def __init__(self, id, ad, email):
        self.id = id
        self.ad = ad
        self.email = email

# Admin erişimi gerektiren fonksiyonlar için decorator
def admin_required(f):
    @functools.wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('admin_giris'):
            flash('Bu sayfa sadece admin erişimine açıktır', 'danger')
            return redirect(url_for('giris'))
        return f(*args, **kwargs)
    return decorated_function

def kullanicilari_oku():
    """kullanicilar.txt dosyasından kullanıcı listesini okur"""
    kullanicilar = []
    
    # Dosya yoksa varsayılan kullanıcılar oluştur
    if not os.path.exists(KULLANICI_DOSYASI):
        varsayilan_kullanicilar = [
            "Tugay Göktürk",
            "Hasan Erinç",
            "Mehmet Öz",
            "Zeynep Kaya",
            "Ahmet Yılmaz"
        ]
        with open(KULLANICI_DOSYASI, 'w', encoding='utf-8') as f:
            for kullanici in varsayilan_kullanicilar:
                f.write(f"{kullanici}\n")
    
    # Farklı kodlamalar deneyerek dosyayı okuma
    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-9', 'cp1254']
    
    for encoding in encodings:
        try:
            with open(KULLANICI_DOSYASI, 'r', encoding=encoding) as f:
                print(f"Kullanıcı dosyası '{encoding}' kodlaması ile başarıyla okundu")
                for i, line in enumerate(f, 1):
                    line = line.strip()
                    if line:
                        kullanicilar.append(Kullanici(str(i), line, ""))
                return kullanicilar
        except UnicodeDecodeError:
            print(f"'{encoding}' kodlaması ile okuma hatası: {str(UnicodeDecodeError)}")
            continue
    
    # Hiçbir kodlama çalışmazsa varsayılan kullanıcılar
    print("Kullanıcı dosyası okunamadı, varsayılan kullanıcılar kullanılıyor")
    kullanicilar = [
        Kullanici("1", "Tugay Göktürk", ""),
        Kullanici("2", "Hasan Erinç", ""),
        Kullanici("3", "Mehmet Öz", ""),
        Kullanici("4", "Zeynep Kaya", ""),
        Kullanici("5", "Ahmet Yılmaz", "")
    ]
    
    return kullanicilar

def kullanici_ekle(ad_soyad):
    """Kullanıcılar dosyasına yeni kullanıcı ekler"""
    try:
        with open(KULLANICI_DOSYASI, 'a', encoding='utf-8') as f:
            f.write(f"{ad_soyad}\n")
        return True
    except Exception as e:
        print(f"Kullanıcı ekleme hatası: {e}")
        return False

def kontrol_excel():
    """Excel dosyasının var olup olmadığını kontrol eder, yoksa oluşturur."""
    try:
        load_workbook(EXCEL_DOSYASI)
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.append([
            'Çeyreklik Dönem', 
            'Vaka Türü', 
            'Kanal', 
            'Ad-Soyad / Üye İş Yeri Adı',
            'Müşteri No / TCKN / Kart No', 
            'Sahte İşlem / Risk Tespiti Tarihi',
            'Kullanılan Tutar', 
            'Müşteri Zararı', 
            'Banka Zararı', 
            'Kurtarılan Tutar',
            'Toplam Risk (TL)', 
            'Vaka Detayı', 
            'Analist Notu', 
            'İşlem Türü',
            'Vaka Giren Kişi', 
            'İnceleyen Kişi',
            'CPP Tespiti', 
            'Pos/ATM', 
            'Kopyalama Onus/Not Onus',
            'Kopyalama Yeri/ Şüpheli İş yeri', 
            'ATM Kart Kopyalama Yöntemi'
        ])
        wb.save(EXCEL_DOSYASI)

def excel_veri_oku():
    """Excel dosyasındaki tüm veriyi okur ve DataFrame olarak döndürür."""
    kontrol_excel()
    wb = load_workbook(EXCEL_DOSYASI)
    sheet = wb.active
    data = sheet.values
    columns = next(data)
    return pd.DataFrame(data, columns=columns)

def excel_veri_yaz(data):
    """Verilen DataFrame'i Excel dosyasına yazar (mevcut veriyi siler)."""
    wb = Workbook()
    ws = wb.active
    for row in dataframe_to_rows(data, header=True, index=False):
        ws.append(row)
    wb.save(EXCEL_DOSYASI)

def excel_veri_ekle(yeni_veri):
    """Yeni bir vakayı Excel dosyasına ekler."""
    kontrol_excel()
    wb = load_workbook(EXCEL_DOSYASI)
    ws = wb.active
    ws.append(list(yeni_veri.values()))
    wb.save(EXCEL_DOSYASI)

def excel_son_kaydi_oku():
    """Excel dosyasındaki son kaydı okur ve DataFrame olarak döndürür."""
    kontrol_excel()
    df = excel_veri_oku()
    if not df.empty:
        return df.iloc[[-1]]  # Son satırı DataFrame olarak döndürür
    return pd.DataFrame()

def excel_kayit_sil(index):
    """Belirtilen indeksteki kaydı Excel dosyasından siler."""
    df = excel_veri_oku()
    if 0 <= index < len(df):
        df = df.drop(df.index[index])
        excel_veri_yaz(df)
        return True
    return False

def excel_kayit_guncelle(index, guncel_veri):
    """Belirtilen indeksteki kaydı verilen verilerle günceller."""
    df = excel_veri_oku()
    if 0 <= index < len(df):
        for key, value in guncel_veri.items():
            if key in df.columns:
                df.at[df.index[index], key] = value
        excel_veri_yaz(df)
        return True
    return False

def risk_hesapla_yeni(musteri_zarari, banka_zarari, kurtarilan_tutar):
    """Yeni toplam risk hesaplar."""
    try:
        musteri_zarari = float(musteri_zarari)
        banka_zarari = float(banka_zarari)
        kurtarilan_tutar = float(kurtarilan_tutar)
        toplam_risk = musteri_zarari + banka_zarari + kurtarilan_tutar
        return toplam_risk
    except ValueError:
        return 0.0

@app.route('/', methods=['GET', 'POST'])
def giris():
    if request.method == 'POST':
        sifre = request.form['sifre']
        
        # Admin şifresi kontrolü
        if 'admin_giris' in request.form and sifre == ADMIN_SIFRE:
            session['giris_yapildi'] = True
            session['admin_giris'] = True
            flash('Admin olarak giriş yaptınız', 'success')
            return redirect(url_for('admin_panel'))
            
        # Normal kullanıcı şifresi kontrolü
        elif sifre == DEFAULT_SIFRE:
            # Kullanıcı seçildi mi kontrol et
            kullanici_id = request.form.get('kullanici_secim')
            if not kullanici_id:
                flash('Lütfen bir kullanıcı seçin!', 'danger')
                kullanicilar = kullanicilari_oku()
                return render_template('giris.html', kullanicilar=kullanicilar, hata='Lütfen bir kullanıcı seçin!')
                
            session['giris_yapildi'] = True
            session['admin_giris'] = False
            session['son_secilen_kullanici_id'] = kullanici_id
            
            # Kullanıcı adını bul ve kaydet
            kullanicilar = kullanicilari_oku()
            for kullanici in kullanicilar:
                if kullanici.id == kullanici_id:
                    session['kullanici_adi'] = kullanici.ad
                    break
                    
            flash('Giriş başarılı', 'success')
            return redirect(url_for('menu'))
        else:
            kullanicilar = kullanicilari_oku()
            flash('Şifre yanlış!', 'danger')
            return render_template('giris.html', kullanicilar=kullanicilar, hata='Şifre yanlış!')
    
    # GET isteği olduğunda giriş formunu göster
    kullanicilar = kullanicilari_oku()
    return render_template('giris.html', kullanicilar=kullanicilar)

@app.route('/logout')
def cikis():
    session.pop('giris_yapildi', None)
    session.pop('admin_giris', None)
    return redirect(url_for('giris'))

def giris_kontrol():
    if not session.get('giris_yapildi'):
        return redirect(url_for('giris'))
    return None

@app.route('/admin')
@admin_required
def admin_panel():
    """Admin paneli ana sayfası"""
    # Kullanıcı sayısı
    total_users = len(kullanicilari_oku())
    
    # E-posta ayarlarını al
    email_ayarlari = email_ayarlarini_kontrol_et()
    
    # Vaka sayısı ve risk tutarları
    try:
        df = pd.read_excel(EXCEL_DOSYASI)
        total_cases = len(df)
        total_risk = "{:,.2f} TL".format(df['Toplam Risk (TL)'].sum())
        
        # Bu ay eklenen vaka sayısı (basit simülasyon)
        # Gerçek uygulamada tarih kontrolü yapılabilir
        month_cases = min(total_cases, int(total_cases * 0.3))
    except:
        total_cases = 0
        total_risk = "0.00 TL"
        month_cases = 0
    
    return render_template('admin_panel.html', 
                          total_users=total_users,
                          total_cases=total_cases,
                          total_risk=total_risk,
                          month_cases=month_cases,
                          email_bildirim_aktif=email_ayarlari["bildirim_aktif"])

@app.route('/son_vaka')
def son_vaka():
    giris_kontrol_sonucu = giris_kontrol()
    if giris_kontrol_sonucu:
        return giris_kontrol_sonucu

    son_vaka_df = excel_son_kaydi_oku()
    vaka = son_vaka_df.to_dict('records')[0] if not son_vaka_df.empty else None
    return render_template('son_vaka.html', vaka=vaka)

@app.route('/sil_son_vaka')
def sil_son_vaka():
    giris_kontrol_sonucu = giris_kontrol()
    if giris_kontrol_sonucu:
        return giris_kontrol_sonucu

    df = excel_veri_oku()
    if not df.empty:
        son_index = len(df) - 1
        if excel_kayit_sil(son_index):
            return redirect(url_for('kayit_ekle'))
        else:
            return "Son vaka silinirken bir hata oluştu."
    else:
        return "Kaydedilmiş vaka bulunmuyor."

@app.route('/duzenle_son_vaka', methods=['GET', 'POST'])
def duzenle_son_vaka():
    """En son eklenen vakayı düzenler."""
    giris_kontrol_sonucu = giris_kontrol()
    if giris_kontrol_sonucu:
        return giris_kontrol_sonucu

    # Verileri oku
    df = excel_veri_oku()
    
    # Boş DataFrame kontrolü
    if df.empty:
        flash('Düzenlenecek vaka bulunamadı!', 'warning')
        return redirect(url_for('kayit_ekle'))
    
    # Son vaka verisini al
    son_vaka = df.iloc[-1].to_dict()
    
    # Kullanıcı listesini al
    kullanicilar = kullanicilari_oku()
    
    # Çeyreklik dönem seçeneklerini al
    ceyreklik_donemler = get_ceyreklik_donemler()
    
    # Yetki kontrolü
    is_admin = session.get('admin_giris', False)
    kullanici_adi = session.get('kullanici_adi', '')
    
    # Admin değilse sadece kendi vakalarını düzenleyebilir
    if not is_admin and son_vaka['Vaka Giren Kişi'] != kullanici_adi:
        flash('Bu vakayı düzenleme yetkiniz yok!', 'danger')
        return redirect(url_for('kayit_ekle'))
    
    if request.method == 'POST':
        ceyrek_donem = request.form['ceyrek_donem']
        vaka_turu = request.form['vaka_turu']
        kanal = request.form['kanal']
        ad_soyad_uye_isyeri = request.form['ad_soyad_uye_isyeri']
        musteri_no_tckn_kart_no = request.form['musteri_no_tckn_kart_no']
        sahte_islem_risk_tarihi = request.form['sahte_islem_risk_tarihi']
        kullanilan_tutar = request.form['kullanilan_tutar']
        musteri_zarari = request.form['musteri_zarari']
        banka_zarari = request.form['banka_zarari']
        kurtarilan_tutar = request.form['kurtarilan_tutar']
        toplam_risk = risk_hesapla_yeni(musteri_zarari, banka_zarari, kurtarilan_tutar)
        
        # Ek bilgileri al
        vaka_detayi = request.form.get('vaka_detayi', '')
        analist_notu = request.form.get('analist_notu', '')
        islem_turu = request.form.get('islem_turu', '')
        
        # Vaka giren kişiyi değiştirmeye izin verme, mevcut değeri kullan
        vaka_giren_kisi = son_vaka['Vaka Giren Kişi']
        
        # İnceleyen kişiyi al
        inceleyen_kisi_id = request.form.get('inceleyen_kisi')
        
        # İnceleyen kişinin adını bul
        if inceleyen_kisi_id in ["Mi4biz", "Şube", "Diğer"]:
            # Özel statik değer, doğrudan kullan
            inceleyen_kisi = inceleyen_kisi_id
        else:
            # Kullanıcılar listesinden bul
            inceleyen_kisi = next((k.ad for k in kullanicilar if k.id == inceleyen_kisi_id), "")

        guncel_kayit = {
            'Çeyreklik Dönem': ceyrek_donem,
            'Vaka Türü': vaka_turu,
            'Kanal': kanal,
            'Ad-Soyad / Üye İş Yeri Adı': ad_soyad_uye_isyeri,
            'Müşteri No / TCKN / Kart No': musteri_no_tckn_kart_no,
            'Sahte İşlem / Risk Tespiti Tarihi': sahte_islem_risk_tarihi,
            'Kullanılan Tutar': kullanilan_tutar,
            'Müşteri Zararı': musteri_zarari,
            'Banka Zararı': banka_zarari,
            'Kurtarılan Tutar': kurtarilan_tutar,
            'Toplam Risk (TL)': toplam_risk,
            'Vaka Detayı': vaka_detayi,
            'Analist Notu': analist_notu,
            'İşlem Türü': islem_turu,
            'Vaka Giren Kişi': vaka_giren_kisi,
            'İnceleyen Kişi': inceleyen_kisi,
        }

        if kanal == 'Kart':
            cpp_tespiti = request.form.get('cpp_tespiti', '')
            if cpp_tespiti and cpp_tespiti != 'Seçiniz':
                guncel_kayit['CPP Tespiti'] = cpp_tespiti
                
            pos_atm = request.form.get('pos_atm', '')
            if pos_atm and pos_atm != 'Seçiniz':
                guncel_kayit['Pos/ATM'] = pos_atm
                
            kopyalama_onus = request.form.get('kopyalama_onus_not_onus', '')
            if kopyalama_onus and kopyalama_onus != 'Seçiniz':
                guncel_kayit['Kopyalama Onus/Not Onus'] = kopyalama_onus
                
            kopyalama_yeri = request.form.get('kopyalama_yeri_supheli_isyeri', '')
            if kopyalama_yeri and kopyalama_yeri != '':
                guncel_kayit['Kopyalama Yeri/ Şüpheli İş yeri'] = kopyalama_yeri
                
            atm_kart_yontemi = request.form.get('atm_kart_kopyalama_yontemi', '')
            if atm_kart_yontemi and atm_kart_yontemi != '':
                guncel_kayit['ATM Kart Kopyalama Yöntemi'] = atm_kart_yontemi

        if excel_kayit_guncelle(df.index[-1], guncel_kayit):
            flash('Vaka başarıyla güncellendi!', 'success')
            return redirect(url_for('kayit_ekle'))
        else:
            flash('Vaka güncellenirken bir hata oluştu!', 'danger')
            return "Son vaka güncellenirken bir hata oluştu."
    
    # Formu göster
    return render_template('duzenle_son_vaka.html', kayit=son_vaka, 
                          kullanicilar=kullanicilar, ceyreklik_donemler=ceyreklik_donemler)

@app.route('/kayit_ekle', methods=['GET', 'POST'])
def kayit_ekle():
    """Yeni bir vaka kaydı oluşturur."""
    kontrol = giris_kontrol()
    if kontrol:
        return kontrol
        
    # Admin erişimi engelle - sadece normal kullanıcılar vaka ekleyebilir
    is_admin = session.get('admin_giris', False)
    if is_admin:
        flash('Admin kullanıcılar vaka ekleyemez. Vaka ekleme işlemi sadece normal kullanıcılar tarafından yapılabilir.', 'warning')
        return redirect(url_for('admin_panel'))

    # Kullanıcı listesini al
    kullanicilar = kullanicilari_oku()
    
    # Çeyreklik dönem seçeneklerini al
    ceyreklik_donemler = get_ceyreklik_donemler()
    
    # Son vaka bilgisini al
    son_vaka_df = excel_son_kaydi_oku()
    son_vaka = son_vaka_df.to_dict('records')[0] if not son_vaka_df.empty else None

    if request.method == 'POST':
        # Temel alanları al
        ceyrek_donem = request.form['ceyrek_donem']
        vaka_turu = request.form['vaka_turu']
        kanal = request.form['kanal']
        ad_soyad_uye_isyeri = request.form['ad_soyad_uye_isyeri']
        musteri_no_tckn_kart_no = request.form['musteri_no_tckn_kart_no']
        sahte_islem_risk_tarihi = request.form['sahte_islem_risk_tarihi']
        kullanilan_tutar = request.form['kullanilan_tutar']
        musteri_zarari = request.form['musteri_zarari']
        banka_zarari = request.form['banka_zarari']
        kurtarilan_tutar = request.form['kurtarilan_tutar']
        toplam_risk = risk_hesapla_yeni(musteri_zarari, banka_zarari, kurtarilan_tutar)
        vaka_detayi = request.form['vaka_detayi']
        analist_notu = request.form['analist_notu']
        islem_turu = request.form['islem_turu']
        
        # Vaka giren kişi olarak sisteme giriş yapan kullanıcıyı kullan
        vaka_giren_kisi = session.get('kullanici_adi', '')
        
        # İnceleyen kişiyi al
        inceleyen_kisi_id = request.form.get('inceleyen_kisi')
        
        # İnceleyen kişinin adını bul
        if inceleyen_kisi_id in ["Mi4biz", "Şube", "Diğer"]:
            # Özel statik değer, doğrudan kullan
            inceleyen_kisi = inceleyen_kisi_id
        else:
            # Kullanıcılar listesinden bul
            inceleyen_kisi = next((k.ad for k in kullanicilar if k.id == inceleyen_kisi_id), "")

        # Tüm kolom bilgilerini önce boş olarak tanımla
        yeni_kayit = {
            'Çeyreklik Dönem': ceyrek_donem,
            'Vaka Türü': vaka_turu,
            'Kanal': kanal,
            'Ad-Soyad / Üye İş Yeri Adı': ad_soyad_uye_isyeri,
            'Müşteri No / TCKN / Kart No': musteri_no_tckn_kart_no,
            'Sahte İşlem / Risk Tespiti Tarihi': sahte_islem_risk_tarihi,
            'Kullanılan Tutar': kullanilan_tutar,
            'Müşteri Zararı': musteri_zarari,
            'Banka Zararı': banka_zarari,
            'Kurtarılan Tutar': kurtarilan_tutar,
            'Toplam Risk (TL)': toplam_risk,
            'Vaka Detayı': vaka_detayi,
            'Analist Notu': analist_notu,
            'İşlem Türü': islem_turu,
            'Vaka Giren Kişi': vaka_giren_kisi,
            'İnceleyen Kişi': inceleyen_kisi,
        }

        # Kart kanalı için ek alanları sadece açıkça belirtilmişse ekle
        if kanal == 'Kart':
            cpp_tespiti = request.form.get('cpp_tespiti', '')
            if cpp_tespiti and cpp_tespiti != 'Seçiniz':
                yeni_kayit['CPP Tespiti'] = cpp_tespiti
                
            pos_atm = request.form.get('pos_atm', '')
            if pos_atm and pos_atm != 'Seçiniz':
                yeni_kayit['Pos/ATM'] = pos_atm
                
            kopyalama_onus = request.form.get('kopyalama_onus_not_onus', '')
            if kopyalama_onus and kopyalama_onus != 'Seçiniz':
                yeni_kayit['Kopyalama Onus/Not Onus'] = kopyalama_onus
                
            kopyalama_yeri = request.form.get('kopyalama_yeri_supheli_isyeri', '')
            if kopyalama_yeri and kopyalama_yeri != '':
                yeni_kayit['Kopyalama Yeri/ Şüpheli İş yeri'] = kopyalama_yeri
                
            atm_kart_yontemi = request.form.get('atm_kart_kopyalama_yontemi', '')
            if atm_kart_yontemi and atm_kart_yontemi != '':
                yeni_kayit['ATM Kart Kopyalama Yöntemi'] = atm_kart_yontemi

        # Excel'e ekle
        excel_veri_ekle(yeni_kayit)
        
        # Yeni eklenen son vaka bilgisini güncelleyelim
        son_vaka_df = excel_son_kaydi_oku()
        son_vaka = son_vaka_df.to_dict('records')[0] if not son_vaka_df.empty else None
        
        # Yeni vakayı günlük listeye ekle
        if son_vaka:
            vaka_gunluk_listeye_ekle(son_vaka)
        
        flash('Kayıt başarıyla eklendi!', 'success')
        return render_template('kayit_formu.html', kullanicilar=kullanicilar, 
                            ceyreklik_donemler=ceyreklik_donemler,
                            onceki_veri=None, son_vaka=son_vaka)
    
    # Formun ilk yüklenişi
    return render_template('kayit_formu.html', kullanicilar=kullanicilar, 
                          ceyreklik_donemler=ceyreklik_donemler,
                          onceki_veri=None, son_vaka=son_vaka)

@app.route('/menu')
def menu():
    giris_kontrol_sonucu = giris_kontrol()
    if giris_kontrol_sonucu:
        return giris_kontrol_sonucu
    return render_template('menu.html')

@app.route('/vakalar')
def vakalar():
    """Vakaları listeler ve filtreleme sağlar."""
    # Giriş kontrolü
    giris_kontrol_sonucu = giris_kontrol()
    if giris_kontrol_sonucu:
        return giris_kontrol_sonucu
    
    # Admin mi kontrolü
    is_admin = session.get('admin_giris', False)
    
    # Filtre parametrelerini al
    secili_kullanici_id = request.args.get('kullanici_id', '')
    secili_kanal = request.args.get('kanal', '')
    secili_ceyrek = request.args.get('ceyrek', '')
    secili_tarih_araligi = request.args.get('tarih_araligi', '')
    
    # Yeni filtre parametreleri
    secili_yil = request.args.get('yil_secim', '')
    secili_ceyrek_kismi = request.args.get('ceyrek_secim', '')
    
    # Excel dosyasından verileri oku
    df = excel_veri_oku()
    
    # Boş DataFrame kontrolü
    if df.empty:
        return render_template('vakalar.html', vakalar=[], is_admin=is_admin, 
                              kullanicilar=[], 
                              secili_kullanici_id=secili_kullanici_id,
                              secili_kanal=secili_kanal, 
                              secili_ceyrek=secili_ceyrek,
                              secili_yil=secili_yil,
                              secili_ceyrek_kismi=secili_ceyrek_kismi,
                              secili_tarih_araligi=secili_tarih_araligi,
                              ceyreklik_donemler=get_ceyreklik_donemler())
    
    # Filtrelemeyi uygula
    filtered_df = df.copy()
    
    # Admin değilse, sadece kendi vakalarını görsün
    if not is_admin:
        kullanici_adi = session.get('kullanici_adi', '')
        filtered_df = filtered_df[filtered_df['Vaka Giren Kişi'] == kullanici_adi]
    elif secili_kullanici_id:  # Admin ve kullanıcı filtresi seçilmişse
        # Seçilen kullanıcının adını bul
        kullanicilar = kullanicilari_oku()
        secili_kullanici_adi = None
        for kullanici in kullanicilar:
            if kullanici.id == secili_kullanici_id:
                secili_kullanici_adi = kullanici.ad
                break
        
        if secili_kullanici_adi:
            filtered_df = filtered_df[filtered_df['Vaka Giren Kişi'] == secili_kullanici_adi]
    
    # Kanal filtrelemesi
    if secili_kanal:
        filtered_df = filtered_df[filtered_df['Kanal'] == secili_kanal]
    
    # Çeyreklik dönem filtrelemesi - birleşik filtre (2025-Q1 gibi)
    if secili_ceyrek:
        filtered_df = filtered_df[filtered_df['Çeyreklik Dönem'] == secili_ceyrek]
    # Sadece yıl filtresi
    elif secili_yil:
        filtered_df = filtered_df[filtered_df['Çeyreklik Dönem'].str.startswith(secili_yil + '-')]
    # Sadece çeyrek kısmı filtresi (Q1, Q2 gibi)
    elif secili_ceyrek_kismi:
        filtered_df = filtered_df[filtered_df['Çeyreklik Dönem'].str.endswith(secili_ceyrek_kismi)]
    
    # Tarih aralığı filtrelemesi
    if secili_tarih_araligi:
        bugun = datetime.datetime.now().date()
        
        if secili_tarih_araligi == 'today':
            # Bugün
            baslangic_tarihi = bugun
        elif secili_tarih_araligi == 'week':
            # Bu hafta (son 7 gün)
            baslangic_tarihi = bugun - datetime.timedelta(days=7)
        elif secili_tarih_araligi == 'month':
            # Bu ay (son 30 gün)
            baslangic_tarihi = bugun - datetime.timedelta(days=30)
        elif secili_tarih_araligi == 'quarter':
            # Bu çeyrek (son 90 gün)
            baslangic_tarihi = bugun - datetime.timedelta(days=90)
        else:
            baslangic_tarihi = None
        
        if baslangic_tarihi:
            # Tarihleri datetime objelerine çevir ve filtrele
            try:
                filtered_df = filtered_df[filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].apply(
                    lambda x: datetime.datetime.strptime(x, '%d-%m-%Y').date() >= baslangic_tarihi
                    if isinstance(x, str) and len(x.strip()) > 0 else False
                )]
            except Exception as e:
                print(f"Tarih filtreleme hatası: {e}")
    
    # Sonuçları liste olarak dönüştür
    vakalar = filtered_df.to_dict('records')
    
    # Kullanıcı listesini al (admin için)
    kullanicilar = kullanicilari_oku() if is_admin else []
    
    return render_template('vakalar.html', vakalar=vakalar, is_admin=is_admin, 
                          kullanicilar=kullanicilar, 
                          secili_kullanici_id=secili_kullanici_id,
                          secili_kanal=secili_kanal, 
                          secili_ceyrek=secili_ceyrek,
                          secili_yil=secili_yil,
                          secili_ceyrek_kismi=secili_ceyrek_kismi,
                          secili_tarih_araligi=secili_tarih_araligi,
                          ceyreklik_donemler=get_ceyreklik_donemler())

@app.route('/kullanici_yonetimi', methods=['GET', 'POST'])
@admin_required
def kullanici_yonetimi():
    """Kullanıcı yönetimi sayfası - sadece admin erişebilir"""
    mesaj = ""
    
    if request.method == 'POST':
        yeni_kullanici = request.form.get('yeni_kullanici', '').strip()
        if yeni_kullanici:
            if kullanici_ekle(yeni_kullanici):
                mesaj = f"'{yeni_kullanici}' başarıyla eklendi."
                flash(f"'{yeni_kullanici}' başarıyla eklendi.", 'success')
            else:
                mesaj = "Kullanıcı eklenirken bir hata oluştu."
                flash("Kullanıcı eklenirken bir hata oluştu.", 'danger')
    
    kullanicilar = kullanicilari_oku()
    return render_template('kullanici_yonetimi.html', 
                          kullanicilar=kullanicilar,
                          mesaj=mesaj)

@app.route('/fraud_trend')
def fraud_trend():
    giris_kontrol_sonucu = giris_kontrol()
    if giris_kontrol_sonucu:
        return giris_kontrol_sonucu
        
    try:
        # Excel dosyasını oku
        df = pd.read_excel(EXCEL_DOSYASI)
        
        # Veri setindeki tüm yılları tespit et
        df['Yıl'] = df['Çeyreklik Dönem'].apply(lambda x: str(x).split('-')[0] if isinstance(x, str) and '-' in str(x) else None)
        available_years = sorted(df['Yıl'].dropna().unique().tolist())
        
        # Varsayılan olarak mevcut yılı seç
        simdiki_yil = str(datetime.datetime.now().year)
        
        # Seçilen yılı request argümanlarından al
        selected_year = request.args.get('selected_year', simdiki_yil)
        
        # Eğer seçilen yıl mevcut değilse ve available_years listesinde yoksa
        # ilk yılı veya mevcut yılı seç
        if not selected_year or selected_year not in available_years:
            # Eğer mevcut yıl liste içinde varsa onu seç, yoksa listenin ilk elemanını seç
            selected_year = simdiki_yil if simdiki_yil in available_years and available_years else available_years[0] if available_years else simdiki_yil
        
        # Seçilen yıla göre verileri filtrele
        filtered_df = df[df['Yıl'] == selected_year]
        year_label = selected_year  # Sayfa başlığı için
        
        # Kanal bazlı analiz için gerekli sütunlar
        channels = ['Dijital Bankacılık', 'Kart', 'POS', 'Başvuru']
        
        # Seçilen yıla göre çeyrek dönemleri belirle
        quarter_ids = [f"{selected_year}-Q1", f"{selected_year}-Q2", f"{selected_year}-Q3", f"{selected_year}-Q4"]
        quarters = quarter_ids + ['Toplam']
        
        # Tutar verisi için veri yapısı oluştur
        tutar_data = {}
        for channel in channels + ['Toplam']:
            tutar_data[channel] = {}
            for quarter in quarters:
                tutar_data[channel][quarter] = {
                    'Risk': 0.0,
                    'Müşteri': 0.0,
                    'Banka': 0.0,
                    'Kurtarılan': 0.0
                }
        
        # Adet verisi için veri yapısı oluştur
        adet_data = {}
        for channel in channels + ['Toplam']:
            adet_data[channel] = {}
            for quarter in quarters:
                adet_data[channel][quarter] = {
                    'Risk': 0,
                    'Müşteri': 0,
                    'Banka': 0,
                    'Kurtarılan': 0
                }
                
        # Her kanal için verileri hesapla
        for channel in channels:
            channel_data = filtered_df[filtered_df['Kanal'] == channel]
            
            # Her çeyrek için hesapla
            for quarter in quarter_ids:
                quarter_data = channel_data[channel_data['Çeyreklik Dönem'] == quarter]
                
                # Tutarlar
                risk_sum = float(quarter_data['Toplam Risk (TL)'].sum())
                musteri_sum = float(quarter_data['Müşteri Zararı'].sum())
                banka_sum = float(quarter_data['Banka Zararı'].sum())
                kurtarilan_sum = float(quarter_data['Kurtarılan Tutar'].sum())
                
                tutar_data[channel][quarter]['Risk'] = risk_sum
                tutar_data[channel][quarter]['Müşteri'] = musteri_sum
                tutar_data[channel][quarter]['Banka'] = banka_sum
                tutar_data[channel][quarter]['Kurtarılan'] = kurtarilan_sum
                
                # Adetler
                risk_count = len(quarter_data)
                musteri_count = len(quarter_data[quarter_data['Müşteri Zararı'] > 0])
                banka_count = len(quarter_data[quarter_data['Banka Zararı'] > 0])
                kurtarilan_count = len(quarter_data[quarter_data['Kurtarılan Tutar'] > 0])
                
                adet_data[channel][quarter]['Risk'] = risk_count
                adet_data[channel][quarter]['Müşteri'] = musteri_count
                adet_data[channel][quarter]['Banka'] = banka_count
                adet_data[channel][quarter]['Kurtarılan'] = kurtarilan_count
                
                # Toplam sütunlara ekle
                for key in ['Risk', 'Müşteri', 'Banka', 'Kurtarılan']:
                    tutar_data['Toplam'][quarter][key] += tutar_data[channel][quarter][key]
                    adet_data['Toplam'][quarter][key] += adet_data[channel][quarter][key]
            
            # Her kanal için toplam değerleri hesapla
            for key in ['Risk', 'Müşteri', 'Banka', 'Kurtarılan']:
                tutar_data[channel]['Toplam'][key] = sum(tutar_data[channel][q][key] for q in quarter_ids)
                adet_data[channel]['Toplam'][key] = sum(adet_data[channel][q][key] for q in quarter_ids)
        
        # Toplam satırı için toplam değerleri hesapla
        for key in ['Risk', 'Müşteri', 'Banka', 'Kurtarılan']:
            tutar_data['Toplam']['Toplam'][key] = sum(tutar_data['Toplam'][q][key] for q in quarter_ids)
            adet_data['Toplam']['Toplam'][key] = sum(adet_data['Toplam'][q][key] for q in quarter_ids)
        
        # Grafik verilerini oluştur
        chart_data = {
            'risk': {},
            'musteri': {},
            'banka': {},
            'kurtarilan': {},
            'years': {}  # Yıllara göre veri
        }
        
        for channel in channels:
            chart_data['risk'][channel] = [
                tutar_data[channel][q]['Risk'] for q in quarter_ids
            ] + [tutar_data[channel]['Toplam']['Risk']]
            
            chart_data['musteri'][channel] = [
                tutar_data[channel][q]['Müşteri'] for q in quarter_ids
            ] + [tutar_data[channel]['Toplam']['Müşteri']]
            
            chart_data['banka'][channel] = [
                tutar_data[channel][q]['Banka'] for q in quarter_ids
            ] + [tutar_data[channel]['Toplam']['Banka']]
            
            chart_data['kurtarilan'][channel] = [
                tutar_data[channel][q]['Kurtarılan'] for q in quarter_ids
            ] + [tutar_data[channel]['Toplam']['Kurtarılan']]
        
        # Toplam değerleri hesapla
        total_risk = tutar_data['Toplam']['Toplam']['Risk']
        musteri_zarari = tutar_data['Toplam']['Toplam']['Müşteri']
        banka_zarari = tutar_data['Toplam']['Toplam']['Banka']
        kurtarilan = tutar_data['Toplam']['Toplam']['Kurtarılan']
        
        # Çeyrek etiketleri
        quarter_labels = ['Q1', 'Q2', 'Q3', 'Q4', 'Toplam']
        
        # Yıllara göre veri yapısı oluştur - tüm yılların verisini hazırla
        # Bu kısım mevcut Yıllara Göre Trend sekmesi için gerekli verileri içeriyor
        year_data = {
            'risk': {},
            'musteri': {},
            'banka': {},
            'kurtarilan': {},
            'adet': {}
        }
        
        # Her kanal ve her yıl için verileri hesapla
        for channel in channels + ['Toplam']:
            year_data['risk'][channel] = []
            year_data['musteri'][channel] = []
            year_data['banka'][channel] = []
            year_data['kurtarilan'][channel] = []
            year_data['adet'][channel] = []
            
            for year in available_years:
                if channel == 'Toplam':
                    year_df = df[df['Yıl'] == year]
                else:
                    year_df = df[(df['Yıl'] == year) & (df['Kanal'] == channel)]
                
                risk_sum = float(year_df['Toplam Risk (TL)'].sum())
                musteri_sum = float(year_df['Müşteri Zararı'].sum())
                banka_sum = float(year_df['Banka Zararı'].sum())
                kurtarilan_sum = float(year_df['Kurtarılan Tutar'].sum())
                count = len(year_df)
                
                year_data['risk'][channel].append(risk_sum)
                year_data['musteri'][channel].append(musteri_sum)
                year_data['banka'][channel].append(banka_sum)
                year_data['kurtarilan'][channel].append(kurtarilan_sum)
                year_data['adet'][channel].append(count)
        
        # "Tüm Zamanlar" toplamı ekle
        for channel in channels + ['Toplam']:
            year_data['risk'][channel].append(sum(year_data['risk'][channel]))
            year_data['musteri'][channel].append(sum(year_data['musteri'][channel]))
            year_data['banka'][channel].append(sum(year_data['banka'][channel]))
            year_data['kurtarilan'][channel].append(sum(year_data['kurtarilan'][channel]))
            year_data['adet'][channel].append(sum(year_data['adet'][channel]))
        
        # Yıl etiketleri (görüntüleme için)
        years_labels = available_years + ['Tüm Zamanlar']
        
        return render_template('fraud_trend.html',
                            chart_data=chart_data,
                            tutar_data=tutar_data,
                            adet_data=adet_data,
                            quarter_labels=quarter_labels,
                            quarter_ids=quarter_ids,
                            total_risk=total_risk,
                            musteri_zarari=musteri_zarari,
                            banka_zarari=banka_zarari,
                            kurtarilan=kurtarilan,
                            year_data=year_data,
                            years=years_labels,
                            selected_year=selected_year,
                            available_years=available_years,
                            year_label=year_label)

    except FileNotFoundError:
        return "Hata: fraud_kayitlari.xlsx dosyası bulunamadı."
    except Exception as e:
        return f"Bir hata oluştu: {e}"

@app.route('/download_excel')
def download_excel():
    """Filtered Excel download based on the same filters used in the vakalar route."""
    try:
        # Giriş kontrolü
        giris_kontrol_sonucu = giris_kontrol()
        if giris_kontrol_sonucu:
            return giris_kontrol_sonucu
            
        # Filtre parametrelerini al
        secili_kullanici_id = request.args.get('kullanici_id', '')
        secili_kanal = request.args.get('kanal', '')
        secili_ceyrek = request.args.get('ceyrek', '')
        secili_tarih_araligi = request.args.get('tarih_araligi', '')
        secili_yil = request.args.get('yil', '')
        secili_ceyrek_kismi = request.args.get('ceyrek_kismi', '')
        search_term = request.args.get('search_term', '').lower() # Yeni parametre
        
        # Excel dosyasından verileri oku
        df = excel_veri_oku()
        
        # Boş DataFrame kontrolü
        if df.empty:
            flash('İndirilecek veri bulunamadı!', 'warning')
            return redirect(url_for('vakalar'))
        
        # Admin mi kontrolü
        is_admin = session.get('admin_giris', False)
        
        # Filtrelemeyi uygula
        filtered_df = df.copy()
        
        # Admin değilse, sadece kendi vakalarını görsün
        if not is_admin:
            kullanici_adi = session.get('kullanici_adi', '')
            filtered_df = filtered_df[filtered_df['Vaka Giren Kişi'] == kullanici_adi]
        elif secili_kullanici_id:  # Admin ve kullanıcı filtresi seçilmişse
            # Seçilen kullanıcının adını bul
            kullanicilar = kullanicilari_oku()
            secili_kullanici_adi = None
            for kullanici in kullanicilar:
                if kullanici.id == secili_kullanici_id:
                    secili_kullanici_adi = kullanici.ad
                    break
            
            if secili_kullanici_adi:
                filtered_df = filtered_df[filtered_df['Vaka Giren Kişi'] == secili_kullanici_adi]
        
        # Kanal filtrelemesi
        if secili_kanal:
            filtered_df = filtered_df[filtered_df['Kanal'] == secili_kanal]
        
        # Çeyreklik dönem filtrelemesi - birleşik filtre (2025-Q1 gibi)
        if secili_ceyrek:
            filtered_df = filtered_df[filtered_df['Çeyreklik Dönem'] == secili_ceyrek]
        # Sadece yıl filtresi
        elif secili_yil:
            filtered_df = filtered_df[filtered_df['Çeyreklik Dönem'].str.startswith(secili_yil + '-')]
        # Sadece çeyrek kısmı filtresi (Q1, Q2 gibi)
        elif secili_ceyrek_kismi:
            filtered_df = filtered_df[filtered_df['Çeyreklik Dönem'].str.endswith(secili_ceyrek_kismi)]
        
        # Tarih aralığı filtrelemesi
        if secili_tarih_araligi:
            bugun = datetime.datetime.now().date()
            
            if secili_tarih_araligi == 'today':
                # Bugün
                baslangic_tarihi = bugun
            elif secili_tarih_araligi == 'week':
                # Bu hafta (son 7 gün)
                baslangic_tarihi = bugun - datetime.timedelta(days=7)
            elif secili_tarih_araligi == 'month':
                # Bu ay (son 30 gün)
                baslangic_tarihi = bugun - datetime.timedelta(days=30)
            elif secili_tarih_araligi == 'quarter':
                # Bu çeyrek (son 90 gün)
                baslangic_tarihi = bugun - datetime.timedelta(days=90)
            else:
                baslangic_tarihi = None
            
            if baslangic_tarihi:
                # Tarihleri datetime objelerine çevir ve filtrele
                try:
                    filtered_df = filtered_df[filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].apply(
                        lambda x: datetime.datetime.strptime(x, '%d-%m-%Y').date() >= baslangic_tarihi
                        if isinstance(x, str) and len(x.strip()) > 0 else False
                    )]
                except Exception as e:
                    print(f"Tarih filtreleme hatası: {e}")
        
        # Arama filtresi (search_term) 
        if search_term:
            # Tüm sütunlarda arama yaparak eşleşenleri al
            mask = filtered_df.apply(lambda row: any(search_term in str(cell).lower() for cell in row), axis=1)
            filtered_df = filtered_df[mask]
        
        # Filtrelenmiş verileri Excel'e yaz
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            filtered_df.to_excel(writer, sheet_name='Filtered_Data', index=False)
            
            # Otomatik sütun genişliklerini ayarla
            worksheet = writer.sheets['Filtered_Data']
            for i, col in enumerate(filtered_df.columns):
                column_width = max(len(col) + 2, filtered_df[col].astype(str).map(len).max() + 2)
                worksheet.set_column(i, i, column_width)
        
        output.seek(0)
        
        # Bugünün tarihini Excel dosya adına ekle
        today = datetime.datetime.now().strftime('%d-%m-%Y')
        filename = f'Fraud_Kayitlari_{today}.xlsx'
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name=filename,
            as_attachment=True
        )
        
    except Exception as e:
        flash(f'Excel indirme sırasında bir hata oluştu: {str(e)}', 'danger')
        return redirect(url_for('vakalar'))

@app.route('/vaka_detay/<int:vaka_index>')
def vaka_detay(vaka_index):
    """Bir vakanın detayını gösterir"""
    if 'giris_yapildi' not in session:
        return redirect(url_for('giris'))
    
    try:
        # Excel dosyasını kontrol et ve oku
        kontrol_excel()
        df = pd.read_excel(EXCEL_DOSYASI, engine='openpyxl')
        
        # Admin tüm kayıtları görebilir, normal kullanıcı sadece kendi vakalarını
        is_admin = session.get('admin_giris', False)
        kullanici_adi = session.get('kullanici_adi', '')
        
        # Tüm veriler
        if not is_admin:
            # Normal kullanıcı sadece kendi vakalarını görebilir
            df = df[df['Vaka Giren Kişi'] == kullanici_adi]
        
        # Eğer vaka bulunamadıysa veya index dışındaysa
        if vaka_index < 0 or vaka_index >= len(df):
            flash('Vaka bulunamadı!', 'danger')
            return redirect(url_for('vakalar'))
        
        # İstenen vakayı al
        vaka = df.iloc[vaka_index].to_dict()
        
        return render_template('vaka_detay.html', vaka=vaka, vaka_index=vaka_index)
    
    except Exception as e:
        return render_template('hata.html', hata=str(e))

@app.route('/duzenle_vaka/<int:vaka_index>', methods=['GET', 'POST'])
def duzenle_vaka(vaka_index):
    """Belirli bir vakayı düzenler"""
    if 'giris_yapildi' not in session:
        return redirect(url_for('giris'))
    
    try:
        kullanicilar = kullanicilari_oku()
        df = excel_veri_oku()
        
        # Çeyreklik dönem seçeneklerini al
        ceyreklik_donemler = get_ceyreklik_donemler()
        
        # Admin tüm kayıtları görebilir, normal kullanıcı sadece kendi vakalarını
        is_admin = session.get('admin_giris', False)
        kullanici_adi = session.get('kullanici_adi', '')
        
        # Filtreleme işlemi
        if not is_admin:
            # Normal kullanıcı sadece kendi vakalarını görebilir
            filtered_df = df[df['Vaka Giren Kişi'] == kullanici_adi]
            if vaka_index >= len(filtered_df):
                flash('Bu vakayı düzenleme yetkiniz yok!', 'danger')
                return redirect(url_for('vakalar'))
            # Gerçek indeksi bul
            real_indices = df[df['Vaka Giren Kişi'] == kullanici_adi].index.tolist()
            if vaka_index >= len(real_indices):
                flash('Vaka bulunamadı!', 'danger')
                return redirect(url_for('vakalar'))
            real_index = real_indices[vaka_index]
        else:
            # Admin için doğrudan indeksi kullan
            if vaka_index < 0 or vaka_index >= len(df):
                flash('Vaka bulunamadı!', 'danger')
                return redirect(url_for('vakalar'))
            real_index = vaka_index
        
        # Düzenlenecek vaka bilgilerini al
        vaka_data = df.iloc[real_index].to_dict()
        
        if request.method == 'POST':
            # Form verilerini al
            ceyrek_donem = request.form.get('ceyrek_donem')
            vaka_turu = request.form.get('vaka_turu')
            kanal = request.form.get('kanal')
            ad_soyad_uye_isyeri = request.form.get('ad_soyad_uye_isyeri')
            musteri_no_tckn_kart_no = request.form.get('musteri_no_tckn_kart_no')
            sahte_islem_risk_tarihi = request.form.get('sahte_islem_risk_tarihi')
            kullanilan_tutar = request.form.get('kullanilan_tutar')
            musteri_zarari = request.form.get('musteri_zarari')
            banka_zarari = request.form.get('banka_zarari')
            kurtarilan_tutar = request.form.get('kurtarilan_tutar')
            vaka_detayi = request.form.get('vaka_detayi')
            analist_notu = request.form.get('analist_notu')
            islem_turu = request.form.get('islem_turu')
            
            # Vaka giren kişiyi değiştirmeye izin verme, mevcut değeri kullan
            vaka_giren_kisi = vaka_data['Vaka Giren Kişi']
            
            # İnceleyen kişiyi al
            inceleyen_kisi_id = request.form.get('inceleyen_kisi')
            
            # İnceleyen kişinin adını bul
            if inceleyen_kisi_id in ["Mi4biz", "Şube", "Diğer"]:
                # Özel statik değer, doğrudan kullan
                inceleyen_kisi = inceleyen_kisi_id
            else:
                # Kullanıcılar listesinden bul
                inceleyen_kisi = next((k.ad for k in kullanicilar if k.id == inceleyen_kisi_id), "")
            
            # Toplam riski hesapla
            try:
                musteri_zarari_float = float(musteri_zarari)
                banka_zarari_float = float(banka_zarari)
                kurtarilan_tutar_float = float(kurtarilan_tutar)
                toplam_risk = risk_hesapla_yeni(musteri_zarari_float, banka_zarari_float, kurtarilan_tutar_float)
            except ValueError:
                toplam_risk = 0.0
            
            # Güncellenmiş veri sözlüğü oluştur
            guncel_kayit = {
                'Çeyreklik Dönem': ceyrek_donem,
                'Vaka Türü': vaka_turu,
                'Kanal': kanal,
                'Ad-Soyad / Üye İş Yeri Adı': ad_soyad_uye_isyeri,
                'Müşteri No / TCKN / Kart No': musteri_no_tckn_kart_no,
                'Sahte İşlem / Risk Tespiti Tarihi': sahte_islem_risk_tarihi,
                'Kullanılan Tutar': kullanilan_tutar,
                'Müşteri Zararı': musteri_zarari,
                'Banka Zararı': banka_zarari,
                'Kurtarılan Tutar': kurtarilan_tutar,
                'Toplam Risk (TL)': toplam_risk,
                'Vaka Detayı': vaka_detayi,
                'Analist Notu': analist_notu,
                'İşlem Türü': islem_turu,
                'Vaka Giren Kişi': vaka_giren_kisi,
                'İnceleyen Kişi': inceleyen_kisi,
            }

            # Kart kanalı için ek alanları ekle
            if kanal == 'Kart':
                cpp_tespiti = request.form.get('cpp_tespiti')
                pos_atm = request.form.get('pos_atm')
                kopyalama_onus_not_onus = request.form.get('kopyalama_onus_not_onus')
                kopyalama_yeri = request.form.get('kopyalama_yeri_supheli_isyeri')
                atm_kart_yontemi = request.form.get('atm_kart_kopyalama_yontemi')
                
                guncel_kayit['CPP Tespiti'] = cpp_tespiti
                guncel_kayit['Pos/ATM'] = pos_atm
                guncel_kayit['Kopyalama Onus/Not Onus'] = kopyalama_onus_not_onus
                guncel_kayit['Kopyalama Yeri/ Şüpheli İş yeri'] = kopyalama_yeri
                guncel_kayit['ATM Kart Kopyalama Yöntemi'] = atm_kart_yontemi

            if excel_kayit_guncelle(real_index, guncel_kayit):
                flash('Vaka başarıyla güncellendi!', 'success')
                return redirect(url_for('vaka_detay', vaka_index=vaka_index))
            else:
                flash('Vaka güncellenirken bir hata oluştu!', 'danger')
        
        # GET isteği veya form doğrulama başarısız olduğunda form görüntülenir
        return render_template('duzenle_vaka.html', vaka=vaka_data, vaka_index=vaka_index, 
                              kullanicilar=kullanicilar, ceyreklik_donemler=ceyreklik_donemler)
    
    except Exception as e:
        flash(f'Bir hata oluştu: {str(e)}', 'danger')
        return redirect(url_for('vakalar'))

@app.route('/sil_vaka/<int:vaka_index>')
def sil_vaka(vaka_index):
    """Belirli bir vakayı siler"""
    if 'giris_yapildi' not in session:
        return redirect(url_for('giris'))
    
    try:
        df = excel_veri_oku()
        
        # Admin tüm kayıtları görebilir, normal kullanıcı sadece kendi vakalarını
        is_admin = session.get('admin_giris', False)
        kullanici_adi = session.get('kullanici_adi', '')
        
        # Filtreleme işlemi
        if not is_admin:
            # Normal kullanıcı sadece kendi vakalarını görebilir
            filtered_df = df[df['Vaka Giren Kişi'] == kullanici_adi]
            if vaka_index >= len(filtered_df):
                flash('Bu vakayı silme yetkiniz yok!', 'danger')
                return redirect(url_for('vakalar'))
            # Gerçek indeksi bul
            real_indices = df[df['Vaka Giren Kişi'] == kullanici_adi].index.tolist()
            if vaka_index >= len(real_indices):
                flash('Vaka bulunamadı!', 'danger')
                return redirect(url_for('vakalar'))
            real_index = real_indices[vaka_index]
        else:
            # Admin için doğrudan indeksi kullan
            if vaka_index < 0 or vaka_index >= len(df):
                flash('Vaka bulunamadı!', 'danger')
                return redirect(url_for('vakalar'))
            real_index = vaka_index
        
        # Vakayı sil
        if excel_kayit_sil(real_index):
            flash('Vaka başarıyla silindi!', 'success')
        else:
            flash('Vaka silinirken bir hata oluştu!', 'danger')
        
        return redirect(url_for('vakalar'))
    
    except Exception as e:
        return render_template('hata.html', hata=str(e))

@app.route('/sil_kullanici/<kullanici_id>')
@admin_required
def sil_kullanici(kullanici_id):
    """Bir kullanıcıyı siler"""
    try:
        # Kullanıcılar listesini oku
        kullanicilar = kullanicilari_oku()
        
        # Silinecek kullanıcıyı bul
        silinecek_kullanici = None
        for kullanici in kullanicilar:
            if kullanici.id == kullanici_id:
                silinecek_kullanici = kullanici
                break
                
        if not silinecek_kullanici:
            flash('Kullanıcı bulunamadı!', 'danger')
            return redirect(url_for('kullanici_yonetimi'))
        
        # Kullanıcının adını kaydet (log mesajı için)
        kullanici_adi = silinecek_kullanici.ad
        
        # Kullanıcıyı listeden çıkar
        yeni_kullanicilar = [k for k in kullanicilar if k.id != kullanici_id]
        
        # Kullanıcılar listesini güncelle
        try:
            with open(KULLANICI_DOSYASI, 'w', encoding='utf-8') as f:
                for kullanici in yeni_kullanicilar:
                    f.write(f"{kullanici.ad}\n")
            
            flash(f'"{kullanici_adi}" kullanıcısı başarıyla silindi.', 'success')
        except Exception as e:
            flash(f'Kullanıcı silinirken bir hata oluştu: {str(e)}', 'danger')
        
        return redirect(url_for('kullanici_yonetimi'))
        
    except Exception as e:
        flash(f'Bir hata oluştu: {str(e)}', 'danger')
        return redirect(url_for('kullanici_yonetimi'))

@app.route('/admin_dashboard')
@admin_required
def admin_dashboard():
    """Admin dashboard sayfası - tarih aralığı bazlı vaka istatistikleri"""
    try:
        # Excel dosyasını kontrol et ve oku
        kontrol_excel()
        df = pd.read_excel(EXCEL_DOSYASI, engine='openpyxl')
        
        # Tarih sütununu datetime formatına dönüştür
        df['Sahte İşlem / Risk Tespiti Tarihi'] = pd.to_datetime(
            df['Sahte İşlem / Risk Tespiti Tarihi'], 
            format='%d/%m/%Y', 
            errors='coerce'
        )
        
        # Filtreleme parametrelerini al
        filter_type = request.args.get('filter_type', 'weekly')  # Varsayılan olarak haftalık
        date_range = request.args.get('date_range', '')
        
        # Bugünün tarihi
        today = datetime.datetime.now().date()
        
        # Filtreleme işlemi
        filtered_df = df.copy()
        
        if filter_type == 'weekly':
            # Bu haftanın başlangıcı (Pazartesi)
            start_of_week = today - timedelta(days=today.weekday())
            filtered_df = filtered_df[filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].dt.date >= start_of_week]
            period_name = f"{start_of_week.strftime('%d/%m/%Y')} - {today.strftime('%d/%m/%Y')}"
            
        elif filter_type == 'monthly':
            # Bu ayın başlangıcı
            start_of_month = today.replace(day=1)
            filtered_df = filtered_df[filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].dt.date >= start_of_month]
            period_name = f"{start_of_month.strftime('%B %Y')}"
            
        elif filter_type == 'yearly':
            # Bu yılın başlangıcı
            start_of_year = today.replace(month=1, day=1)
            filtered_df = filtered_df[filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].dt.date >= start_of_year]
            period_name = f"{start_of_year.year}"
            
        elif filter_type == 'custom' and date_range:
            # Özel tarih aralığı
            try:
                start_date, end_date = date_range.split(' - ')
                start_date = datetime.datetime.strptime(start_date, '%d/%m/%Y').date()
                end_date = datetime.datetime.strptime(end_date, '%d/%m/%Y').date()
                
                filtered_df = filtered_df[
                    (filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].dt.date >= start_date) & 
                    (filtered_df['Sahte İşlem / Risk Tespiti Tarihi'].dt.date <= end_date)
                ]
                period_name = f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}"
            except Exception as e:
                flash(f'Tarih aralığı formatı hatalı: {str(e)}', 'danger')
                filtered_df = df.head(0)  # Boş DataFrame
                period_name = "Geçersiz Tarih Aralığı"
        
        # İndex bilgisini kolonlara ekle
        filtered_df['index'] = filtered_df.index
        
        # İstatistik hesaplamaları
        stats = {
            'toplam_vaka_sayisi': len(filtered_df),
            'toplam_musteri_zarari': filtered_df['Müşteri Zararı'].sum(),
            'toplam_banka_zarari': filtered_df['Banka Zararı'].sum(),
            'toplam_kurtarilan': filtered_df['Kurtarılan Tutar'].sum(),
            'toplam_risk': filtered_df['Toplam Risk (TL)'].sum(),
            'ortalama_risk': filtered_df['Toplam Risk (TL)'].mean() if len(filtered_df) > 0 else 0,
            'period_name': period_name
        }
        
        # DataFrame'i sözlük listesine dönüştür
        vakalar = filtered_df.to_dict('records')
        
        return render_template('admin_dashboard.html', 
                               vakalar=vakalar, 
                               stats=stats, 
                               filter_type=filter_type,
                               date_range=date_range)
                               
    except Exception as e:
        return render_template('hata.html', hata=str(e))

@app.route('/download_full_excel')
def download_full_excel():
    """Directly download the existing fraud_trendi.xlsx file without any filtering."""
    try:
        # Giriş kontrolü
        giris_kontrol_sonucu = giris_kontrol()
        if giris_kontrol_sonucu:
            return giris_kontrol_sonucu
        
        # Admin kontrolü
        if not session.get('admin_giris', False):
            flash('Bu işlem için admin yetkisi gereklidir!', 'danger')
            return redirect(url_for('menu'))
            
        return send_file(EXCEL_DOSYASI, 
                         download_name='fraud_trendi.xlsx',
                         as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    except Exception as e:
        flash(f'Excel indirme sırasında bir hata oluştu: {str(e)}', 'danger')
        return redirect(url_for('admin_panel'))

@app.route('/inceleyenler_raporu')
@admin_required
def inceleyenler_raporu():
    """İnceleyen kişilere göre vaka istatistiklerini gösterir."""
    try:
        # Excel dosyasını kontrol et ve oku
        kontrol_excel()
        df = pd.read_excel(EXCEL_DOSYASI, engine='openpyxl')
        
        # Boş DataFrame kontrolü
        if df.empty:
            flash('Henüz hiç vaka kaydı bulunmamaktadır.', 'warning')
            return redirect(url_for('admin_panel'))
        
        # İnceleyen kişi kolonu yoksa uyarı ver
        if 'İnceleyen Kişi' not in df.columns:
            flash('Excel dosyasında "İnceleyen Kişi" kolonu bulunamadı.', 'warning')
            return redirect(url_for('admin_panel'))
        
        # İnceleyen kişilerin listesini al
        inceleyenler = df['İnceleyen Kişi'].dropna().unique().tolist()
        
        # İnceleyen kişi seçimi
        secili_inceleyen = request.args.get('inceleyen', '')
        if secili_inceleyen and secili_inceleyen not in inceleyenler:
            secili_inceleyen = ''
        
        if not secili_inceleyen and inceleyenler:
            secili_inceleyen = inceleyenler[0]
        
        # İstatistikleri hesapla
        istatistikler = {}
        
        for inceleyen in inceleyenler:
            # İlgili kişinin baktığı vakalar
            inceleyen_df = df[df['İnceleyen Kişi'] == inceleyen]
            
            # Temel istatistikleri hesapla
            toplam_vaka = len(inceleyen_df)
            toplam_musteri_zarari = inceleyen_df['Müşteri Zararı'].sum()
            toplam_banka_zarari = inceleyen_df['Banka Zararı'].sum()
            toplam_kurtarilan = inceleyen_df['Kurtarılan Tutar'].sum()
            toplam_risk = inceleyen_df['Toplam Risk (TL)'].sum()
            
            # Müşteri zararı olan vaka sayısı
            musteri_zararli_vaka = len(inceleyen_df[inceleyen_df['Müşteri Zararı'] > 0])
            
            # Banka zararı olan vaka sayısı
            banka_zararli_vaka = len(inceleyen_df[inceleyen_df['Banka Zararı'] > 0])
            
            # Kanallara göre dağılım
            kanal_dagilimi = inceleyen_df['Kanal'].value_counts().to_dict()
            
            # Çeyreklere göre dağılım
            ceyrek_dagilimi = inceleyen_df['Çeyreklik Dönem'].value_counts().to_dict()
            
            # İstatistikleri sakla
            istatistikler[inceleyen] = {
                'toplam_vaka': toplam_vaka,
                'toplam_musteri_zarari': toplam_musteri_zarari,
                'toplam_banka_zarari': toplam_banka_zarari,
                'toplam_kurtarilan': toplam_kurtarilan,
                'toplam_risk': toplam_risk,
                'musteri_zararli_vaka': musteri_zararli_vaka,
                'banka_zararli_vaka': banka_zararli_vaka,
                'kanal_dagilimi': kanal_dagilimi,
                'ceyrek_dagilimi': ceyrek_dagilimi
            }
        
        # Seçili kişinin vakalarını listelemek için
        secili_vakalar = []
        if secili_inceleyen:
            inceleyen_df = df[df['İnceleyen Kişi'] == secili_inceleyen]
            secili_vakalar = inceleyen_df.to_dict('records')
        
        return render_template('inceleyenler_raporu.html', 
                               inceleyenler=inceleyenler,
                               secili_inceleyen=secili_inceleyen,
                               istatistikler=istatistikler,
                               secili_vakalar=secili_vakalar)
    
    except Exception as e:
        flash(f'Rapor oluşturulurken bir hata oluştu: {str(e)}', 'danger')
        return redirect(url_for('admin_panel'))

@app.route('/fraud_trend_excel/<selected_year>')
def fraud_trend_excel(selected_year):
    """Seçilen yıla göre çeyrek ve kanal bazlı Excel indirme"""
    try:
        # Giriş kontrolü
        giris_kontrol_sonucu = giris_kontrol()
        if giris_kontrol_sonucu:
            return giris_kontrol_sonucu
            
        # Excel dosyasını oku
        df = pd.read_excel(EXCEL_DOSYASI)
        
        # Çeyreklik dönem sütunundan yıl bilgisini çıkar
        df['Yıl'] = df['Çeyreklik Dönem'].apply(lambda x: str(x).split('-')[0] if isinstance(x, str) and '-' in str(x) else None)
        
        # Seçilen yıla göre verileri filtrele
        filtered_df = df[df['Yıl'] == selected_year]
        
        if filtered_df.empty:
            flash('Seçilen yıl için veri bulunamadı!', 'warning')
            return redirect(url_for('fraud_trend'))
        
        # Kanal bazlı analiz için gerekli sütunlar
        channels = ['Dijital Bankacılık', 'Kart', 'POS', 'Başvuru']
        
        # Seçilen yıla göre çeyrek dönemleri belirle
        quarter_ids = [f"{selected_year}-Q1", f"{selected_year}-Q2", f"{selected_year}-Q3", f"{selected_year}-Q4"]
        
        # Bugünün tarihini al ve dosya adına ekle
        today = datetime.datetime.now().strftime('%d-%m-%Y')
        output_filename = f'Fraud_Trend_{selected_year}_{today}.xlsx'
        
        # Geçici dosya oluştur
        output = BytesIO()
        
        # Formatlanmış Excel oluştur
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Tutar analizi worksheet'i oluştur
            tutar_ws = workbook.add_worksheet(f'{selected_year} Kanal Bazlı Tutar Analizi')
            
            # Adet analizi worksheet'i oluştur
            adet_ws = workbook.add_worksheet(f'{selected_year} Kanal Bazlı Adet Analizi')
            
            # Formatlama tanımlamaları
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#2c3e50',
                'color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            q1_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4682b4',
                'color': 'white',
                'align': 'center',
                'valign': 'vcenter', 
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            q2_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#2e8b57',
                'color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            q3_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#cd853f',
                'color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            q4_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#dc143c',
                'color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            total_header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#2f4f4f',
                'color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            subheader_format = workbook.add_format({
                'bold': True,
                'bg_color': '#f5f5f5',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            channel_format = workbook.add_format({
                'bold': True,
                'align': 'left',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            cell_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            money_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '#,##0.00',
                'font_name': 'Arial',
                'font_size': 10
            })
            
            total_row_format = workbook.add_format({
                'bold': True,
                'bg_color': '#f0f0f0',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'font_name': 'Arial',
                'font_size': 10
            })
            
            total_row_money_format = workbook.add_format({
                'bold': True,
                'bg_color': '#f0f0f0',
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'num_format': '#,##0.00',
                'font_name': 'Arial',
                'font_size': 10
            })
            
            # Kolon genişliklerini ayarla (her iki sheet için)
            for sheet in [tutar_ws, adet_ws]:
                sheet.set_column('A:A', 20)  # Kanal sütunu genişliği
                for col_idx in range(1, 21):
                    sheet.set_column(col_idx, col_idx, 15)  # Her veri sütunu 15 birim genişlikte
            
            # Başlık satırları (her iki sheet için)
            headers = [
                # Başlık satırı 1 - Çeyrekler
                ['Kanal',
                 quarter_ids[0], '', '', '',  # Q1
                 quarter_ids[1], '', '', '',  # Q2
                 quarter_ids[2], '', '', '',  # Q3
                 quarter_ids[3], '', '', '',  # Q4
                 f'TOPLAM {selected_year}', '', '', ''  # TOPLAM
                ],
                # Başlık satırı 2 - Sütun başlıkları
                ['',
                 'Risk', 'Müşteri Zararı', 'Banka Zararı', 'Kurtarılan',
                 'Risk', 'Müşteri Zararı', 'Banka Zararı', 'Kurtarılan',
                 'Risk', 'Müşteri Zararı', 'Banka Zararı', 'Kurtarılan',
                 'Risk', 'Müşteri Zararı', 'Banka Zararı', 'Kurtarılan',
                 'Risk', 'Müşteri Zararı', 'Banka Zararı', 'Kurtarılan'
                ]
            ]
            
            # Başlık hücre birleştirmeleri
            for sheet in [tutar_ws, adet_ws]:
                # Başlık satırı 1
                sheet.write(0, 0, 'Kanal', header_format)
                
                # Her çeyrek için hücre birleştirmesi
                sheet.merge_range(0, 1, 0, 4, quarter_ids[0], q1_header_format)
                sheet.merge_range(0, 5, 0, 8, quarter_ids[1], q2_header_format)
                sheet.merge_range(0, 9, 0, 12, quarter_ids[2], q3_header_format)
                sheet.merge_range(0, 13, 0, 16, quarter_ids[3], q4_header_format)
                sheet.merge_range(0, 17, 0, 20, f'TOPLAM {selected_year}', total_header_format)
                
                # Başlık satırı 2
                sheet.write(1, 0, '', header_format)
                
                # Alt başlıklar
                for i in range(1, 21):
                    sheet.write(1, i, headers[1][i], subheader_format)
            
            # Tutar Analizi verilerini hazırla ve yaz
            # TUTAR ANALİZİ İÇİN VERİLERİ HESAPLA
            tutar_data = {}
            
            # Her kanal için verileri hazırla
            for channel in channels:
                tutar_data[channel] = {}
                
                # Her çeyrek için hesaplama yap
                for quarter in quarter_ids:
                    quarter_data = filtered_df[(filtered_df['Kanal'] == channel) & (filtered_df['Çeyreklik Dönem'] == quarter)]
                    
                    risk_sum = float(quarter_data['Toplam Risk (TL)'].sum())
                    musteri_sum = float(quarter_data['Müşteri Zararı'].sum())
                    banka_sum = float(quarter_data['Banka Zararı'].sum())
                    kurtarilan_sum = float(quarter_data['Kurtarılan Tutar'].sum())
                    
                    tutar_data[channel][quarter] = {
                        'Risk': risk_sum,
                        'Müşteri': musteri_sum,
                        'Banka': banka_sum,
                        'Kurtarılan': kurtarilan_sum
                    }
                
                # Bu kanal için yıl toplamını hesapla
                channel_year_data = filtered_df[filtered_df['Kanal'] == channel]
                
                total_risk = float(channel_year_data['Toplam Risk (TL)'].sum())
                total_musteri = float(channel_year_data['Müşteri Zararı'].sum())
                total_banka = float(channel_year_data['Banka Zararı'].sum())
                total_kurtarilan = float(channel_year_data['Kurtarılan Tutar'].sum())
                
                tutar_data[channel]['Toplam'] = {
                    'Risk': total_risk,
                    'Müşteri': total_musteri,
                    'Banka': total_banka,
                    'Kurtarılan': total_kurtarilan
                }
                
            # Toplam satırını hesapla (Tüm kanalların toplamı)
            tutar_data['TOPLAM'] = {}
            
            for quarter in quarter_ids:
                quarter_total = filtered_df[filtered_df['Çeyreklik Dönem'] == quarter]
                
                total_risk = float(quarter_total['Toplam Risk (TL)'].sum())
                total_musteri = float(quarter_total['Müşteri Zararı'].sum())
                total_banka = float(quarter_total['Banka Zararı'].sum())
                total_kurtarilan = float(quarter_total['Kurtarılan Tutar'].sum())
                
                tutar_data['TOPLAM'][quarter] = {
                    'Risk': total_risk,
                    'Müşteri': total_musteri,
                    'Banka': total_banka,
                    'Kurtarılan': total_kurtarilan
                }
            
            # Yıl toplamı
            total_risk = float(filtered_df['Toplam Risk (TL)'].sum())
            total_musteri = float(filtered_df['Müşteri Zararı'].sum())
            total_banka = float(filtered_df['Banka Zararı'].sum())
            total_kurtarilan = float(filtered_df['Kurtarılan Tutar'].sum())
            
            tutar_data['TOPLAM']['Toplam'] = {
                'Risk': total_risk,
                'Müşteri': total_musteri,
                'Banka': total_banka,
                'Kurtarılan': total_kurtarilan
            }
            
            # ADET ANALİZİ İÇİN VERİLERİ HESAPLA
            adet_data = {}
            
            # Her kanal için verileri hazırla
            for channel in channels:
                adet_data[channel] = {}
                
                # Her çeyrek için hesaplama yap
                for quarter in quarter_ids:
                    quarter_data = filtered_df[(filtered_df['Kanal'] == channel) & (filtered_df['Çeyreklik Dönem'] == quarter)]
                    
                    risk_count = len(quarter_data)
                    musteri_count = len(quarter_data[quarter_data['Müşteri Zararı'] > 0])
                    banka_count = len(quarter_data[quarter_data['Banka Zararı'] > 0])
                    kurtarilan_count = len(quarter_data[quarter_data['Kurtarılan Tutar'] > 0])
                    
                    adet_data[channel][quarter] = {
                        'Risk': risk_count,
                        'Müşteri': musteri_count,
                        'Banka': banka_count,
                        'Kurtarılan': kurtarilan_count
                    }
                
                # Bu kanal için yıl toplamını hesapla
                channel_year_data = filtered_df[filtered_df['Kanal'] == channel]
                
                total_risk_count = len(channel_year_data)
                total_musteri_count = len(channel_year_data[channel_year_data['Müşteri Zararı'] > 0])
                total_banka_count = len(channel_year_data[channel_year_data['Banka Zararı'] > 0])
                total_kurtarilan_count = len(channel_year_data[channel_year_data['Kurtarılan Tutar'] > 0])
                
                adet_data[channel]['Toplam'] = {
                    'Risk': total_risk_count,
                    'Müşteri': total_musteri_count,
                    'Banka': total_banka_count,
                    'Kurtarılan': total_kurtarilan_count
                }
                
            # Toplam satırını hesapla (Tüm kanalların toplamı)
            adet_data['TOPLAM'] = {}
            
            for quarter in quarter_ids:
                quarter_total = filtered_df[filtered_df['Çeyreklik Dönem'] == quarter]
                
                total_risk_count = len(quarter_total)
                total_musteri_count = len(quarter_total[quarter_total['Müşteri Zararı'] > 0])
                total_banka_count = len(quarter_total[quarter_total['Banka Zararı'] > 0])
                total_kurtarilan_count = len(quarter_total[quarter_total['Kurtarılan Tutar'] > 0])
                
                adet_data['TOPLAM'][quarter] = {
                    'Risk': total_risk_count,
                    'Müşteri': total_musteri_count,
                    'Banka': total_banka_count,
                    'Kurtarılan': total_kurtarilan_count
                }
            
            # Yıl toplamı
            total_risk_count = len(filtered_df)
            total_musteri_count = len(filtered_df[filtered_df['Müşteri Zararı'] > 0])
            total_banka_count = len(filtered_df[filtered_df['Banka Zararı'] > 0])
            total_kurtarilan_count = len(filtered_df[filtered_df['Kurtarılan Tutar'] > 0])
            
            adet_data['TOPLAM']['Toplam'] = {
                'Risk': total_risk_count,
                'Müşteri': total_musteri_count,
                'Banka': total_banka_count,
                'Kurtarılan': total_kurtarilan_count
            }
            
            # Verileri Excel'e yaz
            row = 2  # İlk veri satırı (0 ve 1 başlıklar için kullanıldı)
            
            # Tutar verileri
            for channel in channels + ['TOPLAM']:
                tutar_ws.write(row, 0, channel, channel_format if channel != 'TOPLAM' else total_row_format)
                
                col_index = 1
                
                # Her çeyrek ve kategorideki verileri yaz
                for quarter in quarter_ids + ['Toplam']:
                    for category in ['Risk', 'Müşteri', 'Banka', 'Kurtarılan']:
                        value = tutar_data[channel][quarter][category]
                        
                        if channel == 'TOPLAM':
                            if value > 0:
                                tutar_ws.write_number(row, col_index, value, total_row_money_format)
                            else:
                                tutar_ws.write(row, col_index, '-', total_row_format)
                        else:
                            if value > 0:
                                tutar_ws.write_number(row, col_index, value, money_format)
                            else:
                                tutar_ws.write(row, col_index, '-', cell_format)
                        
                        col_index += 1
                
                row += 1
            
            # Adet verileri
            row = 2  # İlk veri satırı (0 ve 1 başlıklar için kullanıldı)
            
            for channel in channels + ['TOPLAM']:
                adet_ws.write(row, 0, channel, channel_format if channel != 'TOPLAM' else total_row_format)
                
                col_index = 1
                
                # Her çeyrek ve kategorideki verileri yaz
                for quarter in quarter_ids + ['Toplam']:
                    for category in ['Risk', 'Müşteri', 'Banka', 'Kurtarılan']:
                        value = adet_data[channel][quarter][category]
                        
                        if channel == 'TOPLAM':
                            if value > 0:
                                adet_ws.write_number(row, col_index, value, total_row_format)
                            else:
                                adet_ws.write(row, col_index, '-', total_row_format)
                        else:
                            if value > 0:
                                adet_ws.write_number(row, col_index, value, cell_format)
                            else:
                                adet_ws.write(row, col_index, '-', cell_format)
                        
                        col_index += 1
                
                row += 1
                
        # Dosyayı indir
        output.seek(0)
        
        return send_file(
            output,
            download_name=output_filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    
    except Exception as e:
        flash(f'Excel indirme sırasında bir hata oluştu: {str(e)}', 'danger')
        return redirect(url_for('fraud_trend'))

# Uygulama oluşturma kısmından sonra ekleyin
csrf = CSRFProtect()
csrf.init_app(app)

# Global veri yapısı
son_eklenen_vaka = None
email_zamanlayici = None
gunluk_vakalar = []  # Bir gün içinde eklenen vakaları takip etmek için

# Vaka eklenince e-posta göndermeyi zamanla
def email_zamanlayici_baslat(vaka):
    global son_eklenen_vaka, email_zamanlayici
    
    # Mevcut zamanlayıcıyı iptal et
    if email_zamanlayici is not None:
        email_zamanlayici.cancel()
    
    # Yeni vakayı kaydet
    son_eklenen_vaka = vaka
    
    # E-posta ayarlarını al
    ayarlar = email_ayarlarini_kontrol_et()
    if not ayarlar["bildirim_aktif"]:
        print("E-posta bildirimleri devre dışı. Zamanlayıcı başlatılmadı.")
        return
    
    # Bildirim süresini al (saniye cinsinden)
    bildirim_suresi_saat = ayarlar["bildirim_suresi"]
    bildirim_suresi_saniye = bildirim_suresi_saat * 3600
    
    # Yeni zamanlayıcı başlat
    email_zamanlayici = threading.Timer(bildirim_suresi_saniye, zamanlanmis_email_gonder)
    email_zamanlayici.daemon = True
    email_zamanlayici.start()
    
    print(f"E-posta bildirimi {bildirim_suresi_saat} saat sonrası için zamanlandı.")

# Zamanlanmış e-posta gönder
def zamanlanmis_email_gonder():
    global son_eklenen_vaka
    
    if son_eklenen_vaka is None:
        print("Gönderilecek vaka bulunamadı.")
        return
    
    # E-posta içeriğini hazırla
    konu = f"Yeni Fraud Vaka Bildirimi: {son_eklenen_vaka.get('Vaka Türü', 'Bilinmeyen Vaka')}"
    icerik = vaka_bildirim_email_icerigi_olustur(son_eklenen_vaka)
    
    # E-postayı gönder
    email_gonder(konu, icerik)
    
    # İşlem tamamlandı, son vakayı sıfırla
    son_eklenen_vaka = None

# Flask uygulaması başlatılırken e-posta ayarlarını kontrol et
email_ayarlarini_kontrol_et()

@app.route('/admin/email_yonetimi', methods=['GET', 'POST'])
@admin_required
def email_yonetimi():
    """E-posta bildirim ayarlarını yönetir."""
    
    # Mevcut ayarları al
    email_ayarlari = email_ayarlarini_kontrol_et()
    
    if request.method == 'POST':
        try:
            # Form verilerini al
            yeni_ayarlar = {
                "smtp_sunucu": request.form.get('smtp_sunucu', ''),
                "smtp_port": int(request.form.get('smtp_port', 587)),
                "kullanici_adi": request.form.get('kullanici_adi', ''),
                "gonderen_email": request.form.get('gonderen_email', ''),
                "alici_email": request.form.get('alici_email', ''),
                "bildirim_suresi": int(request.form.get('bildirim_suresi', 3)),
                "bildirim_aktif": request.form.get('bildirim_aktif') == 'on'
            }
            
            # Şifre güncellemesi (boş ise mevcut şifreyi koru)
            yeni_sifre = request.form.get('sifre', '')
            if yeni_sifre:
                yeni_ayarlar["sifre"] = yeni_sifre
            else:
                yeni_ayarlar["sifre"] = email_ayarlari.get("sifre", "")
            
            # Test e-postası gönderme isteği kontrolü
            test_email_gonder = request.form.get('test_email_gonder') == 'on'
            
            # Ayarları güncelle
            if email_ayarlarini_guncelle(yeni_ayarlar):
                flash('E-posta ayarları başarıyla güncellendi!', 'success')
                
                # Test e-postası gönder
                if test_email_gonder:
                    test_baslik = "Fraud Yönetim Sistemi - Test E-postası"
                    test_icerik = """
                    <html>
                    <body style="font-family: Arial, sans-serif; line-height: 1.6;">
                        <div style="max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd;">
                            <h1 style="color: #2c3e50;">E-posta Ayarları Test</h1>
                            <p>Bu bir test e-postasıdır. E-posta bildirimleri doğru şekilde yapılandırılmıştır.</p>
                            <p>Bildirim süresi: {0} saat</p>
                            <p>Tarih/Saat: {1}</p>
                        </div>
                    </body>
                    </html>
                    """.format(
                        yeni_ayarlar["bildirim_suresi"], 
                        datetime.datetime.now().strftime("%d.%m.%Y %H:%M:%S")
                    )
                    
                    if email_gonder(test_baslik, test_icerik):
                        flash('Test e-postası başarıyla gönderildi!', 'success')
                    else:
                        flash('Test e-postası gönderilirken bir hata oluştu. Ayarlarınızı kontrol edin.', 'danger')
            else:
                flash('E-posta ayarları güncellenirken bir hata oluştu!', 'danger')
                
            return redirect(url_for('email_yonetimi'))
        except Exception as e:
            flash(f'Bir hata oluştu: {str(e)}', 'danger')
            print(f"E-posta ayarları güncelleme hatası: {e}")
    
    return render_template('email_yonetimi.html', ayarlar=email_ayarlari)

# Vaka eklenince günlük listeye ekle
def vaka_gunluk_listeye_ekle(vaka):
    global gunluk_vakalar
    gunluk_vakalar.append(vaka)
    print(f"Vaka günlük listeye eklendi. Şu anda listede {len(gunluk_vakalar)} vaka var.")

# Günlük özet e-postası zamanla
def gunluk_ozet_email_zamanlayici_baslat():
    global email_zamanlayici
    
    # Mevcut zamanlayıcıyı iptal et
    if email_zamanlayici is not None:
        email_zamanlayici.cancel()
    
    # E-posta ayarlarını al
    ayarlar = email_ayarlarini_kontrol_et()
    if not ayarlar["bildirim_aktif"]:
        print("E-posta bildirimleri devre dışı. Zamanlayıcı başlatılmadı.")
        return
    
    # Şu anki saat ve dakikayı al
    now = datetime.datetime.now()
    
    # Bir sonraki bildirim zamanını hesapla (saat 8:00)
    bildirim_saati = ayarlar.get("bildirim_saati", 8)  # Varsayılan 8:00
    next_run = now.replace(hour=bildirim_saati, minute=0, second=0, microsecond=0)
    
    # Eğer bu saat geçtiyse, bir sonraki gün için ayarla
    if now >= next_run:
        next_run = next_run + datetime.timedelta(days=1)
    
    # Bir sonraki çalışma zamanına kadar olan farkı hesapla (saniye olarak)
    delay = (next_run - now).total_seconds()
    
    # Yeni zamanlayıcı başlat
    email_zamanlayici = threading.Timer(delay, gunluk_ozet_email_gonder)
    email_zamanlayici.daemon = True
    email_zamanlayici.start()
    
    print(f"Günlük özet e-postası {next_run.strftime('%d.%m.%Y %H:%M')} için zamanlandı.")

# Günlük özet e-postası gönder
def gunluk_ozet_email_gonder():
    global gunluk_vakalar
    
    try:
        # E-posta ayarlarını kontrol et
        ayarlar = email_ayarlarini_kontrol_et()
        if not ayarlar["bildirim_aktif"]:
            print("E-posta bildirimleri devre dışı.")
            return
        
        # Dün eklenmiş tüm vakaları bul
        today = datetime.datetime.now().date()
        yesterday = today - datetime.timedelta(days=1)
        yesterday_str = yesterday.strftime("%d/%m/%Y")
        
        # Excel dosyasından verileri oku
        df = excel_veri_oku()
        
        # Boş DataFrame kontrolü
        if df.empty:
            print("Excel dosyası boş. Gönderilecek vaka bulunamadı.")
            return
        
        # Tarih formatını kontrol et ve dünün vakalarını filtrele
        dun_eklenen_vakalar = []
        
        # Eğer "Sahte İşlem / Risk Tespiti Tarihi" sütunu varsa ve tarih formatında ise
        if "Sahte İşlem / Risk Tespiti Tarihi" in df.columns:
            for index, row in df.iterrows():
                tarih_str = str(row.get("Sahte İşlem / Risk Tespiti Tarihi", ""))
                # Tarih formatı kontrol (GG/AA/YYYY)
                if "/" in tarih_str and tarih_str.count("/") == 2:
                    try:
                        tarih_parts = tarih_str.split("/")
                        if len(tarih_parts) == 3 and len(tarih_parts[0]) == 2 and len(tarih_parts[1]) == 2 and len(tarih_parts[2]) == 4:
                            if tarih_str == yesterday_str:
                                dun_eklenen_vakalar.append(row.to_dict())
                    except:
                        continue
        
        # Dün eklenen vaka yoksa çıkış yap
        if not dun_eklenen_vakalar:
            print("Dün eklenen vaka bulunamadı. Gönderilecek e-posta yok.")
            # Her gün için yeni zamanlayıcı başlat
            gunluk_ozet_email_zamanlayici_baslat()
            return
        
        # E-posta içeriğini oluştur
        now = datetime.datetime.now()
        konu = f"Günlük Fraud Vaka Özeti - {yesterday.strftime('%d.%m.%Y')}"
        
        # HTML içeriği başlat
        icerik = f"""
        <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <div style="max-width: 800px; margin: 0 auto; padding: 20px;">
                <div style="background-color: #2c3e50; color: white; padding: 10px 20px;">
                    <h1 style="color: white; margin: 0;">Günlük Fraud Vaka Özeti</h1>
                    <p style="margin: 5px 0 0 0;">{yesterday.strftime('%d.%m.%Y')} tarihinde eklenen vakalar</p>
                </div>
                <div style="padding: 20px; border: 1px solid #ddd;">
                    <p>Değerli Yetkili,</p>
                    <p>{yesterday.strftime('%d.%m.%Y')} tarihinde toplam <strong>{len(dun_eklenen_vakalar)}</strong> adet yeni vaka eklenmiştir. Aşağıda bu vakaların özeti yer almaktadır:</p>
                    
                    <div style="margin-top: 20px;">
                        <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Eklenen Vakalar</h2>
                        <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
                            <thead>
                                <tr style="background-color: #f5f5f5;">
                                    <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Vaka Türü</th>
                                    <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Kanal</th>
                                    <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Müşteri/Üye İşyeri</th>
                                    <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Toplam Risk (TL)</th>
                                </tr>
                            </thead>
                            <tbody>
        """
        
        # Vakaları tabloya ekle
        for vaka in dun_eklenen_vakalar:
            icerik += f"""
                <tr>
                    <td style="padding: 10px; border: 1px solid #ddd;">{vaka.get('Vaka Türü', '-')}</td>
                    <td style="padding: 10px; border: 1px solid #ddd;">{vaka.get('Kanal', '-')}</td>
                    <td style="padding: 10px; border: 1px solid #ddd;">{vaka.get('Ad-Soyad / Üye İş Yeri Adı', '-')}</td>
                    <td style="padding: 10px; border: 1px solid #ddd;">{vaka.get('Toplam Risk (TL)', '-')} TL</td>
                </tr>
            """
        
        # Özet bilgileri
        toplam_risk = sum(float(vaka.get('Toplam Risk (TL)', 0)) for vaka in dun_eklenen_vakalar)
        icerik += f"""
                            </tbody>
                        </table>
                    </div>
                    
                    <div style="margin-top: 20px;">
                        <h2 style="color: #2c3e50; border-bottom: 1px solid #eee; padding-bottom: 10px;">Özet Bilgiler</h2>
                        <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
                            <tr>
                                <th style="background-color: #f5f5f5; padding: 10px; border: 1px solid #ddd; text-align: left;">Toplam Vaka Sayısı</th>
                                <td style="padding: 10px; border: 1px solid #ddd;">{len(dun_eklenen_vakalar)}</td>
                            </tr>
                            <tr>
                                <th style="background-color: #f5f5f5; padding: 10px; border: 1px solid #ddd; text-align: left;">Toplam Risk</th>
                                <td style="padding: 10px; border: 1px solid #ddd;">{toplam_risk} TL</td>
                            </tr>
                        </table>
                    </div>
                    
                    <p style="margin-top: 20px;">Detaylı bilgi için lütfen <a href="#" style="color: #3498db;">Fraud Yönetim Sistemi</a>'ni ziyaret ediniz.</p>
                    
                    <p style="margin-top: 20px; color: #777; font-size: 0.9em;">Bu e-posta otomatik olarak oluşturulmuştur. Lütfen yanıtlamayınız.</p>
                </div>
                <div style="text-align: center; padding: 10px; color: #777; font-size: 0.8em;">
                    <p>© {now.year} Fraud Yönetim Sistemi</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        # E-postayı gönder
        email_gonder(konu, icerik)
        
        print(f"Günlük özet e-postası başarıyla gönderildi. {len(dun_eklenen_vakalar)} vaka bildirildi.")
        
        # Günlük vakaları sıfırla
        gunluk_vakalar = []
        
        # Her gün için yeni zamanlayıcı başlat
        gunluk_ozet_email_zamanlayici_baslat()
        
    except Exception as e:
        print(f"Günlük özet e-postası gönderilirken hata oluştu: {str(e)}")
        # Hata olsa da yeni zamanlayıcıyı başlat
        gunluk_ozet_email_zamanlayici_baslat()

# Flask uygulaması başlatılırken günlük e-posta zamanlayıcısını başlat
email_ayarlarini_kontrol_et()
gunluk_ozet_email_zamanlayici_baslat()

if __name__ == '__main__':
    app.run(debug=True)