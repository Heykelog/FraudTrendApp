# Fraud Yönetim Sistemi

Bu uygulama, fraud (dolandırıcılık) vakalarının kaydedilmesi, yönetilmesi ve analiz edilmesi için geliştirilmiş web tabanlı bir platformdur.

## Sistem Gereksinimleri

- Python 3.8 veya üzeri
- Web tarayıcı (Chrome, Firefox, Edge önerilir)
- İşletim Sistemi: Windows 10/11, macOS, Linux (Ubuntu 20.04 LTS önerilir)

## Kurulum

### Geliştirme Ortamı (Test)

1. Proje klasörünü bilgisayarınıza indirin
```bash
git clone https://github.com/user/fraud-management-system.git
cd fraud-management-system
```

2. Python sanal ortamını oluşturun ve aktifleştirin
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

3. Gerekli paketleri yükleyin
```bash
pip install -r requirements.txt
```

4. Uygulamayı başlatın
```bash
# Geliştirme modunda
python app.py

# Veya Flask CLI ile
export FLASK_APP=app.py
export FLASK_ENV=development
flask run
```

5. Tarayıcınızda `http://127.0.0.1:5000` adresine giderek uygulamaya erişebilirsiniz

### Üretim Ortamı (Production)

#### Üretim Sunucusu Hazırlığı

1. Gerekli paketleri yükleyin
```bash
pip install -r requirements.txt
pip install gunicorn  # Linux/macOS için
# Windows için Waitress
pip install waitress
```

2. Güvenlik ayarlarını yapın
```bash
# config.py dosyasında SECRET_KEY değerini güvenli bir değerle değiştirin
# ve DEBUG=False olarak ayarlayın
```

3. WSGI sunucusu ile uygulamayı çalıştırın

**Linux/macOS (Gunicorn):**
```bash
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

**Windows (Waitress):**
```bash
waitress-serve --port=8000 app:app
```

4. Reverse Proxy Yapılandırması (Nginx)

```nginx
server {
    listen 80;
    server_name yourdomainname.com;

    location / {
        proxy_pass http://localhost:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

## HTTPS Yapılandırması

### Let's Encrypt ile SSL Sertifikası Alınması

1. Certbot kurun
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install certbot python3-certbot-nginx

# CentOS/RHEL
sudo yum install certbot python3-certbot-nginx
```

2. SSL sertifikası alın
```bash
sudo certbot --nginx -d yourdomainname.com
```

3. Nginx yapılandırmasını güncelleyin
```nginx
server {
    listen 443 ssl;
    server_name yourdomainname.com;
    
    ssl_certificate /etc/letsencrypt/live/yourdomainname.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/yourdomainname.com/privkey.pem;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_prefer_server_ciphers on;
    
    location / {
        proxy_pass http://localhost:8000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}

server {
    listen 80;
    server_name yourdomainname.com;
    return 301 https://$host$request_uri;
}
```

## Veri Yedekleme

Uygulama verileri Excel formatında saklanmaktadır. Düzenli olarak şu dosyaları yedeklemeniz önerilir:

- `vakalar.xlsx` - Tüm vaka kayıtlarını içerir
- `users.xlsx` - Kullanıcı bilgilerini içerir

```bash
# Otomatik yedekleme örneği (Linux/macOS)
crontab -e
# Her gün gece yarısı yedek al
0 0 * * * cp /path/to/vakalar.xlsx /path/to/backups/vakalar_$(date +\%Y\%m\%d).xlsx
0 0 * * * cp /path/to/users.xlsx /path/to/backups/users_$(date +\%Y\%m\%d).xlsx
```

## Güvenlik Önlemleri

1. SECRET_KEY'i güçlü ve benzersiz bir değerle değiştirin
2. DEBUG modunu üretim ortamında kapatın
3. Düzenli olarak güvenlik güncellemelerini kontrol edin
4. Hassas verileri şifreleyerek saklayın
5. IP tabanlı erişim kısıtlamaları ekleyin

## Sorun Giderme

**Uygulama başlatılamıyor:**
- Python sürümünüzün 3.8 veya üstü olduğundan emin olun
- Bağımlılıkların doğru kurulduğunu kontrol edin

**Bağlantı hataları:**
- Firewall ayarlarını kontrol edin
- Port numarasının açık olduğundan emin olun

**Veri görüntüleme sorunları:**
- Excel dosyalarının doğru formatta olduğunu kontrol edin
- Dosya izinlerini kontrol edin

## İletişim ve Destek

Herhangi bir sorun veya öneri için lütfen [e-posta adresi] adresine mail atın veya GitHub üzerinden issue açın.

## Lisans

Bu proje [LİSANS ADI] altında lisanslanmıştır. 