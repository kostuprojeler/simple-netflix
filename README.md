# 🎬 Tkinter Tabanlı Netflix Simülasyonu ve İçerik Yönetim Sistemi

Bu proje, Python'ın yerleşik grafik arayüz kütüphanesi **Tkinter** kullanılarak geliştirilmiş ve ilişkisel bir **MySQL** veritabanı ile entegre çalışan bir Netflix Simülasyonu ve İçerik Yönetim Sistemi uygulamasıdır. Proje; dinamik arama/filtreleme, akıllı içerik öneri algoritması, entegre medya oynatıcı simülasyonu (gerçek zamanlı loglama) ve kapsamlı bir yönetici (admin) paneli barındırmaktadır.

---

## Öne Çıkan Özellikler

### İzleyici (Seyirci) Özellikleri
* **Gelişmiş İçerik Arama:** Yapım adına, türüne (Aksiyon, Komedi vb.) ve içeriğin tipine (Film/Dizi) göre dinamik SQL araması ve puanlama filtresi.
* **Akıllı Karşılama:** Yeni kayıt olan kullanıcılara, seçtikleri 3 favori türe göre veritabanındaki en yüksek puanlı içeriklerin anında listelenmesi.
* **Medya Oynatıcı Simülasyonu:** İzleme dakikasını hafızaya alma, içeriğe 1-10 arası puan verme ve "kaldığın yerden devam et" mekanizması.
* **Gerçek Zamanlı Loglama:** İzlenen her dakikanın ve bölümün veritabanındaki `IzlemeLog` tablosuna kronolojik olarak işlenmesi.
* **Profil ve İstatistik Paneli:** Toplam izleme süresi (dakika), tüketilen yapım sayısı ve puan ortalamasının canlı takibi.

### Yönetici (Admin) Özellikleri
* **Sistem Analitiği Raporları:** En çok izlenen 10 içerik, en yüksek puan alan 10 içerik, en popüler türler, en aktif seyirciler ve son 7 günlük izleme trafiğinin canlı raporlanması.
* **Katalog Yönetimi (CRUD):** Sisteme yeni film/dizi ekleme, mevcut yapımların künyesini/türlerini güncelleme ve katalogdan yapım silme.
* **Kullanıcı Hesap Yönetimi:** Sistemdeki normal üyeleri listeleme, detaylı izleme günlüklerini inceleme ve tek tıkla **Pasif (Ban)** duruma getirme.
* **Güvenli Giriş Bariyeri:** Engellenen (Pasif) kullanıcıların sisteme giriş yapmasını engelleyen güvenlik mimarisi.

---

## Kullanılan Teknolojiler ve Kütüphaneler

* **Programlama Dili:** Python 3.11+
* **Arayüz Motoru:** `tkinter` ve `ttk` (Yerleşik Grafiksel Kullanıcı Arayüzü)
* **Veritabanı Sürücüsü:** `mysql-connector-python` (MySQL Entegrasyonu)
* **Veri Göçü (Migration):** `openpyxl` (Excel dosyasındaki yüzlerce içeriği otomatik okuyup MySQL'e aktarmak için)
* **Ses Motoru:** `pygame.mixer` (Netflix intro jingle sesini oynatmak için)

---

## Veritabanı İlişki Şeması (Database Schema)

Sistem arkada birbiriyle ilişkili toplam **11 adet SQL tablosu** yönetmektedir. `ON DELETE CASCADE` ve `FOREIGN KEY` mimarileri sayesinde veri bütünlüğü en üst düzeyde korunur:

* `Kullanici` & `Rol` (Many-to-One)
* `Program` & `Tur` $\rightarrow$ `ProgramTur` üzerinden (Many-to-Many)
* `Kullanici` & `Tur` $\rightarrow$ `KullaniciTur` üzerinden (Many-to-Many)
* `OturumLog` & `IzlemeLog` (Sistem takip ve analitik logları)

---

## Kurulum ve Çalıştırma

### 1. Gereksinimlerin Yüklenmesi
Bilgisayarınızda Python ve bir MySQL sunucusunun (XAMPP, WampServer vb.) kurulu ve çalışır durumda olduğundan emin olun. Ardından gerekli harici kütüphaneleri terminalden yükleyin:

```bash
pip install mysql-connector-python openpyxl pygame
