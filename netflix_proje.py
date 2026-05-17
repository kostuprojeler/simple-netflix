import mysql.connector
from mysql.connector import Error
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
import time, os, pygame, openpyxl

DB_HOST, DB_USER, DB_PASSWORD, DB_NAME = "localhost", "root", "", "netflix_platformu"

def veritabanini_hazirla():
    try:
        conn = mysql.connector.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD)
        cursor = conn.cursor()
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_NAME} CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;")
        conn.close()
        conn = mysql.connector.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, database=DB_NAME)
        cursor = conn.cursor()

        tablolar = {
            "Rol": "rol_id INT AUTO_INCREMENT PRIMARY KEY, rol_adi VARCHAR(50) UNIQUE NOT NULL",
            "Kullanici": "kullanici_id INT AUTO_INCREMENT PRIMARY KEY, ad VARCHAR(50) NOT NULL, soyad VARCHAR(50) NOT NULL, email VARCHAR(100) UNIQUE NOT NULL, sifre VARCHAR(100) NOT NULL, dogum_tarihi DATE NOT NULL, cinsiyet VARCHAR(20), ulke VARCHAR(50), rol_id INT, durum VARCHAR(20) DEFAULT 'Aktif', FOREIGN KEY (rol_id) REFERENCES Rol(rol_id) ON DELETE SET NULL",
            "Program": "program_id INT AUTO_INCREMENT PRIMARY KEY, program_adi VARCHAR(150) NOT NULL, aciklama TEXT, program_tipi VARCHAR(20) NOT NULL, bolum_sayisi INT DEFAULT 1, program_uzunlugu INT, yayin_yili INT",
            "Tur": "tur_id INT AUTO_INCREMENT PRIMARY KEY, tur_adi VARCHAR(50) UNIQUE NOT NULL",
            "ProgramTur": "program_id INT, tur_id INT, PRIMARY KEY (program_id, tur_id), FOREIGN KEY (program_id) REFERENCES Program(program_id) ON DELETE CASCADE, FOREIGN KEY (tur_id) REFERENCES Tur(tur_id) ON DELETE CASCADE",
            "KullaniciTur": "kullanici_id INT, tur_id INT, PRIMARY KEY (kullanici_id, tur_id), FOREIGN KEY (kullanici_id) REFERENCES Kullanici(kullanici_id) ON DELETE CASCADE, FOREIGN KEY (tur_id) REFERENCES Tur(tur_id) ON DELETE CASCADE",
            "KullaniciProgram": "kullanici_id INT, program_id INT, izleme_tarihi DATE, izlenen_bolum INT DEFAULT 1, izleme_suresi INT DEFAULT 0, puan INT NULL, tamamlandi_mi TINYINT DEFAULT 0, PRIMARY KEY (kullanici_id, program_id), FOREIGN KEY (kullanici_id) REFERENCES Kullanici(kullanici_id) ON DELETE CASCADE, FOREIGN KEY (program_id) REFERENCES Program(program_id) ON DELETE CASCADE",
            "Favori": "kullanici_id INT, program_id INT, PRIMARY KEY (kullanici_id, program_id), FOREIGN KEY (kullanici_id) REFERENCES Kullanici(kullanici_id) ON DELETE CASCADE, FOREIGN KEY (program_id) REFERENCES Program(program_id) ON DELETE CASCADE",
            "Bolum": "bolum_id INT AUTO_INCREMENT PRIMARY KEY, program_id INT, bolum_no INT NOT NULL, bolum_adi VARCHAR(100), uzunluk INT DEFAULT 45, FOREIGN KEY (program_id) REFERENCES Program(program_id) ON DELETE CASCADE",
            "OturumLog": "log_id INT AUTO_INCREMENT PRIMARY KEY, kullanici_id INT, giris_tarihi DATETIME NOT NULL, FOREIGN KEY (kullanici_id) REFERENCES Kullanici(kullanici_id) ON DELETE CASCADE",
            "IzlemeLog": "izleme_id INT AUTO_INCREMENT PRIMARY KEY, kullanici_id INT, program_id INT, izleme_tarihi DATE NOT NULL, izlenen_bolum INT DEFAULT 1, izleme_suresi INT DEFAULT 0, tamamlandi_mi TINYINT DEFAULT 0, FOREIGN KEY (kullanici_id) REFERENCES Kullanici(kullanici_id) ON DELETE CASCADE, FOREIGN KEY (program_id) REFERENCES Program(program_id) ON DELETE CASCADE"
        }
        for t_ad, t_sql in tablolar.items():
            cursor.execute(f"CREATE TABLE IF NOT EXISTS {t_ad} ({t_sql}) ENGINE=InnoDB;")

        cursor.execute("INSERT IGNORE INTO Rol (rol_id, rol_adi) VALUES (1, 'Normal'), (2, 'Yonetici');")
        turler = ['Aksiyon ve Macera', 'Komedi', 'Drama', 'Bilim Kurgu ve Fantastik', 'Belgesel', 'Çocuk ve Aile', 'Romantik', 'Gerilim', 'Korku']
        for idx, t_ad in enumerate(turler, 1):
            cursor.execute("INSERT IGNORE INTO Tur (tur_id, tur_adi) VALUES (%s, %s);", (idx, t_ad))

        cursor.execute("SELECT COUNT(*) FROM Program")
        if cursor.fetchone()[0] == 0:
            excel_yolu = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Netflix DB (1).xlsx")
            if os.path.exists(excel_yolu):
                sheet = openpyxl.load_workbook(excel_yolu, data_only=True).active
                for auto_id, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=1):
                    if not row[0]: continue
                    p_adi, t_ham, p_tipi = str(row[0]).strip(), str(row[1]).strip(), str(row[2]).strip()
                    b_sayi = int(row[3]) if len(row) > 3 and row[3] is not None else 1
                    uzunluk = int(row[4]) if len(row) > 4 and row[4] is not None else 0
                    yil = int(row[5]) if len(row) > 5 and row[5] is not None else 2026
                    ozet = str(row[6]).strip() if len(row) > 6 and row[6] else f"{p_adi} icerigi."
                    cursor.execute("INSERT INTO Program VALUES (%s, %s, %s, %s, %s, %s, %s)", (auto_id, p_adi, ozet, p_tipi, b_sayi, uzunluk, yil))
                    if p_tipi == "Dizi":
                        for b_no in range(1, b_sayi + 1):
                            cursor.execute("INSERT INTO Bolum (program_id, bolum_no, bolum_adi, uzunluk) VALUES (%s, %s, %s, %s)", (auto_id, b_no, f"{b_no}. Bölüm", uzunluk))
                    for t_adi in [t.strip() for t in t_ham.split(",") if t.strip()]:
                        cursor.execute("INSERT IGNORE INTO Tur (tur_adi) VALUES (%s)", (t_adi,))
                        cursor.execute("SELECT tur_id FROM Tur WHERE tur_adi=%s", (t_adi,))
                        t_id = cursor.fetchone()[0]
                        cursor.execute("INSERT IGNORE INTO ProgramTur VALUES (%s, %s)", (auto_id, t_id))
            else:
                print(f"Uyarı: Excel dosyası bulunamadı, varsayılan kurulum yapılıyor.")

        cursor.execute("""
            REPLACE INTO Kullanici (kullanici_id, ad, soyad, email, sifre, dogum_tarihi, rol_id, durum) 
            VALUES (1, 'Admin', 'Sistem', 'batinn@netflix.com', '123456', '1990-01-01', 2, 'Aktif');
        """)
        conn.commit(); conn.close()
    except Error as e: print(f"MySQL Başlatma Hatası: {e}")

def get_db_connection(): return mysql.connector.connect(host=DB_HOST, user=DB_USER, password=DB_PASSWORD, database=DB_NAME)

veritabanini_hazirla()

class NetflixApp(tk.Tk):
    def __init__(self):
        super().__init__() #mmiras alma yapısı
        self.title("Netflix beta")
        self.geometry("1020x740")
        self.configure(bg="#0a0f1d")
        pygame.mixer.init()
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("Treeview", bg="#101726", fg="#ffffff", fieldbackground="#101726", rowheight=30, borderwidth=0)
        self.style.map("Treeview", background=[("selected", "#00d2ff")], foreground=[("selected", "#0a0f1d")])
        self.style.configure("Treeview.Heading", bg="#1a233a", fg="#ff9f1c", font=("Arial", 10, "bold"), borderwidth=0)
        self.style.configure("TCombobox", fieldbackground="#101726", background="#1a233a", foreground="white")
        self.aktif_kullanici = None ##kapsülleme yapısı
        self.sayfa_degistir(GirisSayfasi)

    def sayfa_degistir(self, sayfa_sinifi):
        yeni_sayfa = sayfa_sinifi(self)
        yeni_sayfa.configure(bg="#0a0f1d")
        if hasattr(self, "mevcut_sayfa"): self.mevcut_sayfa.destroy()
        self.mevcut_sayfa = yeni_sayfa
        self.mevcut_sayfa.pack(fill="both", expand=True)

    def netflix_intro_oynat(self, sonraki_sayfa):
        ses = os.path.join(os.path.dirname(os.path.abspath(__file__)), "nouveau-jingle-netflix.mp3")
        try:
            pygame.mixer.quit()
            pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=512)
            if os.path.exists(ses):
                pygame.mixer.music.load(ses)
                pygame.mixer.music.play()
        except: pass
        if hasattr(self, "mevcut_sayfa"): self.mevcut_sayfa.destroy()
        ifr = tk.Frame(self, bg="#000000"); ifr.pack(fill="both", expand=True); self.mevcut_sayfa = ifr
        cv = tk.Canvas(ifr, width=400, height=300, bg="#000000", highlightthickness=0); cv.place(relx=0.5, rely=0.5, anchor="center")
        for i in range(1, 40):
            cv.delete("all")
            cv.create_rectangle(150, 150 - i*3, 175, 150 + i*3, fill="#E50914", outline="")
            cv.create_rectangle(225, 150 - i*3, 250, 150 + i*3, fill="#E50914", outline="")
            if i > 10: cv.create_polygon(150, 150 - i*3, 175, 150 - i*3, 250, 150 + i*3, 225, 150 + i*3, fill="#b80710", outline="")
            self.update(); time.sleep(0.015)
        time.sleep(1.5); self.sayfa_degistir(sonraki_sayfa)

class GirisSayfasi(tk.Frame):
    def __init__(self, master):
        super().__init__(master) #miras alma yapısı
        cf = tk.Frame(self, bg="#101726", padx=40, pady=40); cf.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(cf, text="NETFLIX", font=("Trebuchet MS", 28, "bold", "italic"), fg="#E50914", bg="#101726").pack(pady=(0, 20))
        tk.Label(cf, text="Oturum Aç", font=("Arial", 16, "bold"), fg="white", bg="#101726").pack(anchor="w", pady=(0, 15))
        tk.Label(cf, text="E-mail", fg="#00d2ff", bg="#101726").pack(anchor="w")
        self.ent_email = tk.Entry(cf, width=30, bg="#1a233a", fg="white", insertbackground="white", bd=0, font=("Arial", 11)); self.ent_email.pack(ipady=8, pady=(5, 15))  #kapsülleme 
        tk.Label(cf, text="Şifre", fg="#00d2ff", bg="#101726").pack(anchor="w")
        self.ent_sifre = tk.Entry(cf, show="*", width=30, bg="#1a233a", fg="white", insertbackground="white", bd=0, font=("Arial", 11)); self.ent_sifre.pack(ipady=8, pady=(5, 20))
        tk.Button(cf, text="Giriş Yap", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 11, "bold"), bd=0, command=self.giris_yap).pack(fill="x", ipady=8, pady=(0, 10))
        tk.Button(cf, text="Kayıt Ol", bg="#1a233a", fg="white", font=("Arial", 10), bd=0, command=lambda: master.sayfa_degistir(KayitSayfasi)).pack(fill="x", ipady=6)

    def giris_yap(self):
        email, sifre = self.ent_email.get().strip(), self.ent_sifre.get()
        if not email or not sifre: return messagebox.showwarning("Hata", "Boş alan bırakmayınız!")
        if "@" not in email or "." not in email: return messagebox.showwarning("Hata", "Hatalı e-mail formatı!")
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT * FROM Kullanici WHERE email = %s AND sifre = %s", (email, sifre))
            user = cursor.fetchone()
            if user:
                if user[9] == 'Pasif': return messagebox.showerror("Engellendi", "Hesabınız pasiftir!"), conn.close()
                cursor.execute("INSERT INTO OturumLog (kullanici_id, giris_tarihi) VALUES (%s, %s)", (user[0], datetime.now()))
                conn.commit(); conn.close(); self.master.aktif_kullanici = user
                self.master.netflix_intro_oynat(YoneticiPaneli if user[8] == 2 else KullaniciAnaSayfasi)
            else:
                conn.close(); messagebox.showerror("Hata", "Yanlış şifre veya e-mail!")
        except Error as e: messagebox.showerror("Hata", str(e))

class KayitSayfasi(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        cf = tk.Frame(self, bg="#101726", padx=30, pady=20); cf.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(cf, text="Hesap Oluştur", font=("Arial", 16, "bold"), fg="white", bg="#101726").pack(pady=(0, 15))
        ff = tk.Frame(cf, bg="#101726"); ff.pack()
        fields = [("Ad:", "ad"), ("Soyad:", "soyad"), ("E-mail:", "email"), ("Şifre:", "sifre"), ("Şifre Tekrar:", "sifre2"), ("Doğum Tarihi(YIL-AY-GÜN):", "dt"), ("Cinsiyet:", "cin"), ("Ülke:", "ulke")]
        self.entries = {}
        for idx, (txt, key) in enumerate(fields):
            tk.Label(ff, text=txt, fg="#00d2ff", bg="#101726", font=("Arial", 9)).grid(row=idx//2, column=(idx%2)*2, sticky="w", padx=5)
            ent = tk.Entry(ff, show="*" if "sifre" in key else "", width=18, bg="#1a233a", fg="white", bd=0, insertbackground="white", font=("Arial", 10))
            ent.grid(row=idx//2, column=(idx%2)*2+1, ipady=5, pady=5, padx=5); self.entries[key] = ent
        tk.Label(cf, text="3 Farklı Favori Tür Seçiniz (Ctrl ile çoklu):", font=("Arial", 9, "bold"), fg="#ff9f1c", bg="#101726").pack(pady=(10, 5))
        self.lst_turler = tk.Listbox(cf, selectmode="multiple", height=4, width=38, bg="#1a233a", fg="white", bd=0, selectbackground="#00d2ff")
        self.lst_turler.pack()
        try:
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT tur_id, tur_adi FROM Tur")
            for r in cursor.fetchall(): self.lst_turler.insert(tk.END, f"{r[0]}- {r[1]}")
            conn.close()
        except: pass
        tk.Button(cf, text="Kaydı Tamamla", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 11, "bold"), bd=0, command=self.kaydet).pack(fill="x", ipady=6, pady=(15, 5))
        tk.Button(cf, text="Geri Dön", bg="#1a233a", fg="white", bd=0, command=lambda: master.sayfa_degistir(GirisSayfasi)).pack(fill="x", ipady=4)

    def kaydet(self):
        d = {k: v.get().strip() for k, v in self.entries.items()}
        sel = self.lst_turler.curselection()
        if any(not val for val in d.values()) or len(sel) != 3: return messagebox.showwarning("Hata", "Boş alan bırakmayınız ve tam 3 favori tür seçiniz!")
        if d["sifre"] != d["sifre2"]: return messagebox.showwarning("Hata", "Şifreler uyuşmuyor!")
        if len(d["sifre"]) < 6: return messagebox.showwarning("Hata", "Şifre en az 6 karakter olmalıdır!")
        try:
            if datetime.strptime(d["dt"], "%Y-%m-%d") > datetime.now(): return messagebox.showwarning("Hata", "Geçersiz doğum tarihi!")
        except: return messagebox.showwarning("Hata", "Doğum tarihi YYYY-MM-DD olmalıdır!")
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM Kullanici WHERE email=%s", (d["email"],))
            if cursor.fetchone(): return messagebox.showwarning("Hata", "Bu e-mail kayıtlı!"), conn.close()
            cursor.execute("INSERT INTO Kullanici (ad, soyad, email, sifre, dogum_tarihi, cinsiyet, ulke, rol_id) VALUES (%s,%s,%s,%s,%s,%s,%s,1)", (d["ad"], d["soyad"], d["email"], d["sifre"], d["dt"], d["cin"], d["ulke"]))
            k_id = cursor.lastrowid
            sec_turler = []
            for idx in sel:
                t_v = self.lst_turler.get(idx).split("- ")
                sec_turler.append(t_v[1]); cursor.execute("INSERT INTO KullaniciTur VALUES (%s, %s)", (k_id, int(t_v[0])))
            conn.commit()
            msg = "Kaydınız Başarılı! En yüksek puanlı içerikleriniz:\n\n"
            for t_ad in sec_turler:
                cursor.execute("SELECT p.program_id, p.program_adi, IFNULL(ROUND(AVG(kp.puan),1), 'Yok') FROM Program p JOIN ProgramTur pt ON p.program_id=pt.program_id JOIN Tur t ON pt.tur_id=t.tur_id LEFT JOIN KullaniciProgram kp ON p.program_id=kp.program_id WHERE t.tur_adi=%s GROUP BY p.program_id ORDER BY AVG(kp.puan) DESC LIMIT 2", (t_ad,))
                msg += f"Kategori: {t_ad}:\n"
                for row in cursor.fetchall(): msg += f"  - {row[1]} (Puan: {row[2]})\n"
            conn.close(); messagebox.showinfo("Hoş Geldiniz!", msg); self.master.sayfa_degistir(GirisSayfasi)
        except Error as e: messagebox.showerror("Hata", str(e))

class KullaniciAnaSayfasi(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.k_id = master.aktif_kullanici[0]
        nv = tk.Frame(self, bg="#101726", padx=15, pady=10); nv.pack(fill="x")
        tk.Label(nv, text="NETFLIX", font=("Trebuchet MS", 22, "bold", "italic"), fg="#E50914", bg="#101726").pack(side="left", padx=(0, 20))
        self.lbl_seyirci_adi = tk.Label(nv, text=f"Seyirci: {master.aktif_kullanici[1]} {master.aktif_kullanici[2]}", fg="white", bg="#101726", font=("Arial", 11, "bold")); self.lbl_seyirci_adi.pack(side="left")
        tk.Button(nv, text="Çıkış", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 9, "bold"), bd=0, padx=10, command=self.cikis_yap).pack(side="right", padx=5)
        for t, cmd in [("Profil Ayarlarım", self.profil_penceresi), ("Favorilerim", self.favori_penceresi), ("Geçmişim", self.gecmis_penceresi)]:
            tk.Button(nv, text=t, bg="#1a233a", fg="white", font=("Arial", 9), bd=0, padx=10, command=cmd).pack(side="right", padx=5)
        
        fl = tk.LabelFrame(self, text=" Gelişmiş İçerik Arama ve Filtreleme ", font=("Arial", 10, "bold"), bg="#101726", fg="white", bd=0, padx=15, pady=15); fl.pack(fill="x", padx=15, pady=10)
        
        tk.Label(fl, text="Yapım Adı:", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).grid(row=0, column=0, padx=5, sticky="w")
        self.ent_ad = tk.Entry(fl, width=15, bg="#1a233a", fg="white", bd=0, insertbackground="white", font=("Arial", 10)); self.ent_ad.grid(row=0, column=1, padx=5, ipady=4)
        
        tk.Label(fl, text="Tip:", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).grid(row=0, column=2, padx=5, sticky="w")
        self.cmb_tip = ttk.Combobox(fl, values=["Hepsi", "Film", "Dizi"], width=8, font=("Arial", 10)); self.cmb_tip.set("Hepsi"); self.cmb_tip.grid(row=0, column=3, padx=5)
        
        tk.Label(fl, text="Tür:", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).grid(row=0, column=4, padx=5, sticky="w")
        self.cmb_tur = ttk.Combobox(fl, width=16, font=("Arial", 10)); self.cmb_tur.grid(row=0, column=5, padx=5); self.turler_yukle()
        
        tk.Label(fl, text="Yıl:", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).grid(row=0, column=6, padx=5, sticky="w")
        self.ent_yil = tk.Entry(fl, width=8, bg="#1a233a", fg="white", bd=0, insertbackground="white", font=("Arial", 10)); self.ent_yil.insert(0, "Örn: 2026"); self.ent_yil.grid(row=0, column=7, padx=5, ipady=4)
        
        tk.Label(fl, text="Min Puan:", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).grid(row=0, column=8, padx=5, sticky="w")
        self.ent_min_puan = tk.Entry(fl, width=8, bg="#1a233a", fg="white", bd=0, insertbackground="white", font=("Arial", 10)); self.ent_min_puan.insert(0, "1-10"); self.ent_min_puan.grid(row=0, column=9, padx=5, ipady=4)
        
        tk.Button(fl, text="Filtrele", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 10, "bold"), bd=0, padx=15, cursor="hand2", command=self.icerik_listele_ve_filtrele).grid(row=0, column=10, padx=15, ipady=4)
        
        tc = tk.Frame(self, bg="#0a0f1d"); tc.pack(fill="both", expand=True, padx=15, pady=5)
        cols = ("ID", "Yapım Adı", "Tip", "Türler", "Yıl", "Bölüm", "Uzunluk", "Ort. Puan", "İzlenme")
        self.tree = ttk.Treeview(tc, columns=cols, show="headings")
        for c in cols: self.tree.heading(c, text=c); self.tree.column(c, width=95, anchor="center")
        self.tree.pack(fill="both", expand=True); self.tree.bind("<Double-1>", self.detaya_git)
        self.oneri_box = tk.LabelFrame(self, text=" Sana Özel Akıllı Öneriler ", font=("Arial", 10, "bold"), bg="#101726", fg="#00d2ff", bd=0, padx=10, pady=10); self.oneri_box.pack(fill="x", padx=15, pady=15)
        self.icerik_listele_ve_filtrele(); self.onerileri_olustur()

    def ana_sayfa_ismini_tazele(self): self.lbl_seyirci_adi.config(text=f"Seyirci: {self.master.aktif_kullanici[1]} {self.master.aktif_kullanici[2]}")
    def turler_yukle(self):
        try:
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT tur_adi FROM Tur")
            self.cmb_tur["values"] = ["Hepsi"] + [r[0] for r in cursor.fetchall()]; self.cmb_tur.set("Hepsi"); conn.close()
        except: pass
    def cikis_yap(self): self.master.aktif_kullanici = None; self.master.sayfa_degistir(GirisSayfasi)
    def icerik_listele_ve_filtrele(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            q = "SELECT p.program_id, p.program_adi, p.program_tipi, (SELECT GROUP_CONCAT(t.tur_adi SEPARATOR ', ') FROM ProgramTur pt JOIN Tur t ON pt.tur_id = t.tur_id WHERE pt.program_id = p.program_id), p.yayin_yili, p.bolum_sayisi, p.program_uzunlugu, IFNULL(ROUND(AVG(kp.puan), 1), 'Puan Yok') as ort_puan, COUNT(kp.izleme_tarihi) FROM Program p LEFT JOIN KullaniciProgram kp ON p.program_id = kp.program_id WHERE 1=1"
            p = []
            if self.ent_ad.get(): q += " AND p.program_adi LIKE %s"; p.append(f"%{self.ent_ad.get()}%")
            if self.cmb_tip.get() != "Hepsi": q += " AND p.program_tipi = %s"; p.append(self.cmb_tip.get())
            if self.cmb_tur.get() != "Hepsi": q += " AND p.program_id IN (SELECT pt.program_id FROM ProgramTur pt JOIN Tur t ON pt.tur_id = t.tur_id WHERE t.tur_adi = %s)"; p.append(self.cmb_tur.get())
            yil_val = self.ent_yil.get().strip()
            if yil_val and "Örn" not in yil_val: q += " AND p.yayin_yili = %s"; p.append(int(yil_val))
            q += " GROUP BY p.program_id"
            puan_val = self.ent_min_puan.get().strip()
            if puan_val and "-" not in puan_val: q += " HAVING ort_puan >= %s AND ort_puan != 'Puan Yok'"; p.append(float(puan_val))
            cursor.execute(q, p)
            for r in cursor.fetchall(): self.tree.insert("", "end", values=r)
            conn.close()
        except Exception as e: print(e)
    def onerileri_olustur(self):
        for w in self.oneri_box.winfo_children(): w.destroy()
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT p.program_id, p.program_adi FROM Program p JOIN ProgramTur pt ON p.program_id = pt.program_id WHERE pt.tur_id IN (SELECT tur_id FROM KullaniciTur WHERE kullanici_id = %s) LIMIT 6", (self.k_id,))
            for r in cursor.fetchall():
                lbl = tk.Label(self.oneri_box, text=f"Yapım: {r[1]}", fg="white", bg="#1a233a", font=("Arial", 10, "bold"), padx=10, pady=6); lbl.pack(side="left", padx=8, pady=5)
                lbl.bind("<Enter>", lambda e, l=lbl: l.configure(bg="#00d2ff", fg="#0a0f1d"))
                lbl.bind("<Leave>", lambda e, l=lbl: l.configure(bg="#1a233a", fg="white"))
                lbl.bind("<Button-1>", lambda e, pid=r[0]: IcerikDetaySayfasi(self, pid))
            conn.close()
        except: pass
    def detaya_git(self, event):
        if self.tree.selection(): IcerikDetaySayfasi(self, self.tree.item(self.tree.selection()[0], "values")[0])
    def profil_penceresi(self): ProfilSayfasi(self)
    def favori_penceresi(self): FavorilerSayfasi(self)
    def gecmis_penceresi(self): IzlemeGecmisiSayfasi(self)

class IcerikDetaySayfasi(tk.Toplevel):
    def __init__(self, master, program_id):
        super().__init__(master)
        if master.master.aktif_kullanici[8] == 2: 
            messagebox.showerror("Hata", "Yönetici izleyemez!")
            self.destroy()
            return
        self.p_id, self.k_id = program_id, master.k_id
        self.geometry("520x500"); self.configure(bg="#101726"); self.title("İçerik Bilgi Kartı")
        self.detay_goster()

    def detay_goster(self):
        for w in self.winfo_children(): w.destroy()
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT * FROM Program WHERE program_id=%s", (self.p_id,))
            p = cursor.fetchone()
            cursor.execute("SELECT izlenen_bolum, izleme_suresi, puan FROM KullaniciProgram WHERE kullanici_id=%s AND program_id=%s", (self.k_id, self.p_id))
            log = cursor.fetchone()
            cursor.execute("SELECT 1 FROM Favori WHERE kullanici_id=%s AND program_id=%s", (self.k_id, self.p_id))
            fav = cursor.fetchone()
            cursor.execute("SELECT GROUP_CONCAT(t.tur_adi SEPARATOR ', ') FROM ProgramTur pt JOIN Tur t ON pt.tur_id=t.tur_id WHERE pt.program_id=%s", (self.p_id,))
            t_str = cursor.fetchone()[0] or "Yok"
            conn.close()
            tk.Label(self, text=p[1], font=("Arial", 18, "bold"), fg="#ff9f1c", bg="#101726").pack(pady=10)
            bx = tk.Frame(self, bg="#1a233a", padx=15, pady=15); bx.pack(fill="x", padx=20)
            tk.Label(bx, text=f"Özet: {p[2]}", wraplength=440, justify="left", fg="#ccc", bg="#1a233a", font=("Arial", 10, "italic")).pack(anchor="w", pady=5)
            tk.Label(bx, text=f"Tür Bilgisi: {t_str}", fg="#00d2ff", bg="#1a233a", font=("Arial", 10, "bold")).pack(anchor="w")
            tk.Label(bx, text=f"Tip: {p[3]}  |  Yıl: {p[6]}  |  Bölüm Sayısı: {p[4]}", fg="white", bg="#1a233a").pack(anchor="w", pady=5)
            
            if p[3] == "Dizi":
                tk.Label(self, text="Bölüm Seçin:", fg="white", bg="#101726", font=("Arial", 9, "bold")).pack(pady=(10, 2))
                self.cmb_bolumler = ttk.Combobox(self, values=[f"{i}. Bölüm" for i in range(1, p[4] + 1)], width=15)
                self.cmb_bolumler.set(f"{log[0] if log else 1}. Bölüm"); self.cmb_bolumler.pack()
            if log:
                tk.Label(self, text=f"Puanınız: {log[2] if log[2] else 'Yok'} | Konum: {log[0]}. Bölüm {log[1]}. Dk", fg="#00d2ff", bg="#101726", font=("Arial", 10, "bold")).pack(pady=5)
            tk.Button(self, text="Kaldığın Yerden Devam Et" if log else "Şimdi İzle", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 10, "bold"), bd=0, command=self.izle).pack(pady=10, ipady=5, fill="x", padx=20)
            tk.Button(self, text="Favorilerden Çıkar" if fav else "Favorilerime Ekle", bg="#1a233a", fg="white", bd=0, command=self.favori_islem).pack(fill="x", padx=20, ipady=4)
        except Error as e: print(e)

    def favori_islem(self):
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM Favori WHERE kullanici_id=%s AND program_id=%s", (self.k_id, self.p_id))
            if cursor.fetchone(): cursor.execute("DELETE FROM Favori WHERE kullanici_id=%s AND program_id=%s", (self.k_id, self.p_id))
            else: cursor.execute("INSERT INTO Favori VALUES (%s, %s)", (self.k_id, self.p_id))
            conn.commit(); conn.close(); self.detay_goster()
        except: pass
    def izle(self): IzlemeEkrani(self, self.p_id, int(self.cmb_bolumler.get().split(".")[0]) if hasattr(self, 'cmb_bolumler') else 1)

class IzlemeEkrani(tk.Toplevel):
    def __init__(self, detay_sayfasi, program_id, secili_bolum):
        super().__init__(detay_sayfasi)
        self.ds, self.p_id, self.k_id, self.secili_b = detay_sayfasi, program_id, detay_sayfasi.k_id, secili_bolum
        self.geometry("380x360"); self.configure(bg="#0a0f1d"); self.title("Medya Oynatıcı")
        self.init_s = 0
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT izleme_suresi FROM KullaniciProgram WHERE kullanici_id=%s AND program_id=%s AND izlenen_bolum=%s", (self.k_id, self.p_id, self.secili_b))
            log = cursor.fetchone()
            if log and log[0] > 0:
                if messagebox.askyesno("Devam Et", f"Bölümün {log[0]}. dakikasında kaldınız. Devam edilsin mi?"): self.init_s = log[0]
            conn.close()
        except: pass
        tk.Label(self, text="MEDYA OYNATICI SİMÜLASYONU", font=("Arial", 12, "bold"), fg="#ff9f1c", bg="#0a0f1d").pack(pady=15)
        tk.Label(self, text=f"İzlenen Bölüm: {self.secili_b}", fg="white", bg="#0a0f1d").pack()
        tk.Label(self, text="Dakika Konumu:", fg="#00d2ff", bg="#0a0f1d").pack(pady=(10,0))
        self.ent_s = tk.Entry(self, bg="#101726", fg="white", bd=0, justify="center", insertbackground="white"); self.ent_s.insert(0, str(self.init_s)); self.ent_s.pack(ipady=5, pady=5)
        tk.Label(self, text="Puan Ver (1 - 10):", fg="#00d2ff", bg="#0a0f1d").pack(pady=5)
        self.ent_p = tk.Entry(self, width=8, bg="#101726", fg="white", bd=0, justify="center", insertbackground="white"); self.ent_p.pack(ipady=5)
        tk.Button(self, text="Hafızaya Al ve Kapat", bg="#1a233a", fg="white", bd=0, command=lambda: self.kaydet(0)).pack(fill="x", padx=30, ipady=5, pady=(20, 5))
        tk.Button(self, text="İzlemeyi Tamamen Bitir", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 10, "bold"), bd=0, command=lambda: self.kaydet(1)).pack(fill="x", padx=30, ipady=5)

    def kaydet(self, durum):
        try:
            s, p_val = int(self.ent_s.get()), self.ent_p.get().strip()
            puan = int(p_val) if p_val else None
            if puan and (puan < 1 or puan > 10): return messagebox.showwarning("Hata", "Puan 1-10 kalmalıdır!")
        except: return messagebox.showwarning("Hata", "Sayısal giriş yapınız!")
        
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT program_tipi, program_uzunlugu, bolum_sayisi FROM Program WHERE program_id=%s", (self.p_id,))
            prog_info = cursor.fetchone()
            
            log_suresi = s
            kaydedilecek_bolum = self.secili_b
            kaydedilecek_dakika = s
            tamam_bayragi = durum
            
            if durum == 1:
                log_suresi = prog_info[1] if prog_info[1] > 0 else 45
                kaydedilecek_dakika = 0  
                
                if prog_info[0] == "Dizi":
                    if self.secili_b < prog_info[2]:
                        kaydedilecek_bolum = self.secili_b + 1
                        tamam_bayragi = 0  
                    else:
                        tamam_bayragi = 1  
                else:
                    tamam_bayragi = 1  
            
            t_bugun = datetime.now().strftime("%Y-%m-%d")
            cursor.execute("INSERT INTO KullaniciProgram VALUES (%s,%s,%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE izleme_tarihi=VALUES(izleme_tarihi), izlenen_bolum=VALUES(izlenen_bolum), izleme_suresi=VALUES(izleme_suresi), puan=IFNULL(VALUES(puan), puan), tamamlandi_mi=VALUES(tamamlandi_mi)", (self.k_id, self.p_id, t_bugun, kaydedilecek_bolum, kaydedilecek_dakika, puan, tamam_bayragi))
            cursor.execute("INSERT INTO IzlemeLog (kullanici_id, program_id, izleme_tarihi, izlenen_bolum, izleme_suresi, tamamlandi_mi) VALUES (%s,%s,%s,%s,%s,%s)", (self.k_id, self.p_id, t_bugun, self.secili_b, log_suresi, durum))
            conn.commit(); conn.close(); messagebox.showinfo("Başarılı", "İzleme durumu hafızaya işlendi."); self.ds.detay_goster(); self.ds.master.icerik_listele_ve_filtrele(); self.destroy()
        except Error as e: messagebox.showerror("Hata", str(e))

class ProfilSayfasi(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.master_page, self.k_id = master, master.k_id
        self.geometry("460x650"); self.configure(bg="#101726"); self.title("Profil Ayarlarım")
        self.profil_arayuzunu_kur()

    def _form_satiri(self, txt, val, secret=False):
        rf = tk.Frame(self, bg="#101726"); rf.pack(fill="x", padx=35, pady=4)
        tk.Label(rf, text=txt, fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold"), width=22, anchor="w").pack(side="left")
        ent = tk.Entry(rf, bg="#1a233a", fg="white", bd=0, font=("Arial", 10), insertbackground="white", show="*" if secret else "")
        ent.insert(0, val); ent.pack(side="right", fill="x", expand=True, ipady=5)
        return ent

    def profil_arayuzunu_kur(self):
        for w in self.winfo_children(): w.destroy()
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT ad, soyad, email, sifre, ulke, DATE_FORMAT(dogum_tarihi, '%Y-%m-%d') FROM Kullanici WHERE kullanici_id=%s", (self.k_id,))
            u = cursor.fetchone()
            cursor.execute("SELECT GROUP_CONCAT(t.tur_adi SEPARATOR ', ') FROM KullaniciTur kt JOIN Tur t ON kt.tur_id = t.tur_id WHERE kt.kullanici_id=%s", (self.k_id,))
            f_tur = cursor.fetchone()[0] or "Seçilmemiş"
            cursor.execute("SELECT SUM(izleme_suresi) FROM IzlemeLog WHERE kullanici_id=%s", (self.k_id,))
            sure = cursor.fetchone()[0] or 0
            cursor.execute("SELECT COUNT(DISTINCT program_id) FROM IzlemeLog WHERE kullanici_id=%s", (self.k_id,))
            sayi = cursor.fetchone()[0]
            cursor.execute("SELECT IFNULL(ROUND(AVG(puan), 1), '0.0') FROM KullaniciProgram WHERE kullanici_id=%s AND puan IS NOT NULL", (self.k_id,))
            ort = cursor.fetchone()[0]; conn.close()

            tk.Label(self, text="HESAP AYARLARI VE GÜNCELLEME", font=("Arial", 13, "bold"), fg="#ff9f1c", bg="#101726").pack(pady=(15, 10))
            self.inputs = {
                "ad": self._form_satiri("Adınız:", u[0]),
                "soyad": self._form_satiri("Soyadınız:", u[1]),
                "email": self._form_satiri("E-mail Adresi:", u[2]),
                "dt": self._form_satiri("Doğum Tarihi (YYYY-MM-DD):", u[5]),
                "ulke": self._form_satiri("Ülke:", u[4] or ""),
                "sifre": self._form_satiri("Şifre Güncelleme:", u[3], secret=True)
            }
            tk.Button(self, text="Değişiklikleri Kaydet", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 10, "bold"), bd=0, command=self.profil_guncelle_db).pack(fill="x", padx=35, ipady=7, pady=15)
            sb = tk.LabelFrame(self, text=" Platform İstatistikleriniz ", font=("Arial", 10, "bold"), bg="#1a233a", fg="white", bd=0, padx=15, pady=15); sb.pack(fill="both", expand=True, padx=35, pady=(0, 20))
            tk.Label(sb, text=f"Favoriler: {f_tur}", fg="#00d2ff", bg="#1a233a", font=("Arial", 9, "bold"), wraplength=350, justify="left").pack(anchor="w", pady=3)
            tk.Label(sb, text=f"İzleme Süresi: {sure} Dakika", fg="white", bg="#1a233a").pack(anchor="w", pady=3)
            tk.Label(sb, text=f"Yapım Sayısı: {sayi} Adet", fg="white", bg="#1a233a").pack(anchor="w", pady=3)
            tk.Label(sb, text=f"Puan Ortalamanız: {ort} / 10", fg="#ff9f1c", bg="#1a233a", font=("Arial", 10, "bold")).pack(anchor="w", pady=3)
        except Error as e: messagebox.showerror("Hata", str(e))

    def profil_guncelle_db(self):
        v = {k: ent.get().strip() for k, ent in self.inputs.items()}
        if any(not val for val in v.values()) or len(v["sifre"]) < 6 or "@" not in v["email"] or "." not in v["email"]:
            return messagebox.showwarning("Hata", "Girdilerinizi kontrol ediniz.")
        try:
            if datetime.strptime(v["dt"], "%Y-%m-%d") > datetime.now(): return messagebox.showwarning("Hata", "Geçersiz Tarih!")
        except: return messagebox.showwarning("Hata", "Tarih YYYY-MM-DD olmalı!")
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT kullanici_id FROM Kullanici WHERE email=%s AND kullanici_id != %s", (v["email"], self.k_id))
            if cursor.fetchone(): return messagebox.showwarning("Hata", "Bu e-mail kullanımda!"), conn.close()
            cursor.execute("UPDATE Kullanici SET ad=%s, soyad=%s, email=%s, dogum_tarihi=%s, ulke=%s, sifre=%s WHERE kullanici_id=%s", (v["ad"], v["soyad"], v["email"], v["dt"], v["ulke"], v["sifre"], self.k_id))
            conn.commit()
            cursor.execute("SELECT * FROM Kullanici WHERE kullanici_id=%s", (self.k_id,))
            self.master_page.master.aktif_kullanici = cursor.fetchone()
            conn.close(); self.master_page.ana_sayfa_ismini_tazele(); messagebox.showinfo("Başarılı", "Profil güncellendi."); self.destroy()
        except Error as e: messagebox.showerror("Hata", str(e))

class FavorilerSayfasi(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.k_id = master.k_id
        self.geometry("360x400"); self.configure(bg="#101726"); self.title("Favori Listesi")
        tk.Label(self, text="BEĞENDİĞİM İÇERİKLER", font=("Arial", 12, "bold"), fg="white", bg="#101726").pack(pady=10)
        ff = tk.Frame(self, bg="#101726"); ff.pack(fill="x", padx=20)
        tk.Label(ff, text="Filtre:", fg="#00d2ff", bg="#101726").pack(side="left")
        self.cmb_fav_tur = ttk.Combobox(ff, width=15); self.cmb_fav_tur.pack(side="left", padx=5)
        tk.Button(ff, text="Uygula", bg="#ff9f1c", bd=0, command=self.favori_listele).pack(side="left")
        self.lb = tk.Listbox(self, bg="#1a233a", fg="white", bd=0, selectbackground="#00d2ff", font=("Arial", 11)); self.lb.pack(fill="both", expand=True, padx=20, pady=10)
        self.turleri_yukle(); self.favori_listele()

    def turleri_yukle(self):
        try:
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT tur_adi FROM Tur")
            self.cmb_fav_tur["values"] = ["Hepsi"] + [r[0] for r in cursor.fetchall()]; self.cmb_fav_tur.set("Hepsi"); conn.close()
        except: pass
    def favori_listele(self):
        self.lb.delete(0, tk.END)
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            q = "SELECT p.program_id, p.program_adi FROM Favori f JOIN Program p ON f.program_id = p.program_id WHERE f.kullanici_id=%s"
            p = [self.k_id]
            if self.cmb_fav_tur.get() != "Hepsi" and self.cmb_fav_tur.get() != "":
                q += " AND p.program_id IN (SELECT pt.program_id FROM ProgramTur pt JOIN Tur t ON pt.tur_id = t.tur_id WHERE t.tur_adi = %s)"; p.append(self.cmb_fav_tur.get())
            cursor.execute(q, p)
            for r in cursor.fetchall(): self.lb.insert(tk.END, f"ID: {r[0]} | {r[1]}")
            conn.close()
        except: pass

class IzlemeGecmisiSayfasi(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.geometry("540x340"); self.configure(bg="#101726"); self.title("Gerçek Zamanlı İzleme Günlüğü")
        tk.Label(self, text="SİSTEM İZLEME GEÇMİŞİNİZ (LOG)", font=("Arial", 12, "bold"), fg="white", bg="#101726").pack(pady=10)
        self.tree_history = ttk.Treeview(self, columns=("İçerik", "Tarih", "Bölüm", "Süre", "Durum"), show="headings")
        for c in ("İçerik", "Tarih", "Bölüm", "Süre", "Durum"): self.tree_history.heading(c, text=c); self.tree_history.column(c, width=100, anchor="center")
        self.tree_history.pack(fill="both", expand=True, padx=15, pady=10)
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT p.program_adi, il.izleme_tarihi, il.izlenen_bolum, il.izleme_suresi, CASE WHEN il.tamamlandi_mi = 1 THEN 'Bitti' ELSE 'Kaldı' END FROM IzlemeLog il JOIN Program p ON il.program_id = p.program_id WHERE il.kullanici_id=%s ORDER BY il.izleme_id DESC", (master.k_id,))
            for r in cursor.fetchall(): self.tree_history.insert("", "end", values=r)
            conn.close()
        except: pass

class YoneticiPaneli(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        nv = tk.Frame(self, bg="#101726", padx=15, pady=10); nv.pack(fill="x")
        tk.Label(nv, text="NETFLIX PANEL (ADMIN)", font=("Impact", 20, "bold"), fg="#ff9f1c", bg="#101726").pack(side="left")
        tk.Button(nv, text="Çıkış", bg="#ff4d4d", fg="white", font=("Arial", 10, "bold"), bd=0, padx=12, command=self.hesap_degistir).pack(side="right", padx=5)
        for t, cmd, bg_c in [("Seçili Yapımı Sil", self.secili_yapimi_sil, "#ff4d4d"), ("Yeni Yapım Ekle", self.yeni_icerik_ekle, "#ff9f1c"), ("Tür Yönetimi", self.tur_yonetim_penceresi, "#ff9f1c"), ("Kullanıcıları Yönet", self.kullanici_yonetim_penceresi, "#00d2ff"), ("Yenile / Raporla", self.raporla, "#1a233a")]:
            tk.Button(nv, text=t, bg=bg_c, fg="black" if bg_c!="#1a233a" and bg_c!="#ff4d4d" else "white", font=("Arial", 10, "bold" if bg_c!="#1a233a" else "normal"), bd=0, padx=12, command=cmd).pack(side="right", padx=5)
        tc = tk.Frame(self, bg="#0a0f1d"); tc.pack(fill="x", padx=15, pady=5)
        self.tree_admin = ttk.Treeview(tc, columns=("ID", "Yapım Adı", "Tip", "Yıl", "Bölüm Sayısı"), show="headings", height=5)
        for c in ("ID", "Yapım Adı", "Tip", "Yıl", "Bölüm Sayısı"): self.tree_admin.heading(c, text=c); self.tree_admin.column(c, width=150, anchor="center")
        self.tree_admin.pack(fill="x", expand=True); self.tree_admin.bind("<Double-1>", self.yapim_guncelleme_penceresi)
        tk.Label(self, text=" Sistem Analitik ve İstatistik Rapor Logları ", font=("Arial", 12, "bold"), fg="white", bg="#0a0f1d").pack(pady=(10, 5))
        self.txt_panel = tk.Text(self, height=15, width=105, bg="#101726", fg="#00d2ff", insertbackground="white", bd=0, font=("Consolas", 10), padx=10, pady=10); self.txt_panel.pack(pady=5)
        self.raporla()

    def friendship_forward(self):
        pass

    def hesap_degistir(self): self.master.aktif_kullanici = None; self.master.sayfa_degistir(GirisSayfasi)
    
    def yapim_guncelleme_penceresi(self, event):
        if not self.tree_admin.selection(): return
        p_id = self.tree_admin.item(self.tree_admin.selection()[0], "values")[0]
        
        conn = get_db_connection(); cursor = conn.cursor()
        cursor.execute("SELECT program_adi, aciklama, program_tipi, bolum_sayisi, program_uzunlugu, yayin_yili FROM Program WHERE program_id=%s", (p_id,))
        p_data = cursor.fetchone()
        conn.close()

        fm = tk.Toplevel(self); fm.geometry("360x520"); fm.configure(bg="#101726"); fm.title("Yapım Kataloğunu Düzenle")
        
        flds = ["Yapım Adı:", "Açıklama:", "Uzunluk (Dk):", "Yayın Yılı:", "Bölüm Sayısı:"]
        defaults = [p_data[0], p_data[1], p_data[4], p_data[5], p_data[3]]
        entries = {}
        
        for idx, lbl in enumerate(flds):
            tk.Label(fm, text=lbl, fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).pack(pady=(5,0))
            e = tk.Entry(fm, width=30, bg="#1a233a", fg="white", bd=0, insertbackground="white")
            e.insert(0, str(defaults[idx])); e.pack(ipady=4, pady=2)
            entries[lbl] = e

        tk.Label(fm, text="Kategori Türleri (Çoklu Seçim İçin Ctrl):", fg="#ff9f1c", bg="#101726", font=("Arial", 9, "bold")).pack(pady=(5,0))
        lst_t = tk.Listbox(fm, selectmode="multiple", height=4, width=30, bg="#1a233a", fg="white", bd=0, selectbackground="#00d2ff")
        lst_t.pack(pady=4)
        
        try:
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT tur_id, tur_adi FROM Tur")
            for r in cursor.fetchall(): lst_t.insert(tk.END, f"{r[0]}- {r[1]}")
            conn.close()
        except: pass

        def db_g():
            try:
                conn = get_db_connection(); cursor = conn.cursor()
                cursor.execute("UPDATE Program SET program_adi=%s, aciklama=%s, program_uzunlugu=%s, yayin_yili=%s, bolum_sayisi=%s WHERE program_id=%s", 
                               (entries["Yapım Adı:"].get(), entries["Açıklama:"].get(), int(entries["Uzunluk (Dk):"].get()), int(entries["Yayın Yılı:"].get()), int(entries["Bölüm Sayısı:"].get()), p_id))
                
                sel_t = lst_t.curselection()
                if sel_t:
                    cursor.execute("DELETE FROM ProgramTur WHERE program_id=%s", (p_id,))
                    for idx in sel_t:
                        t_id = int(lst_t.get(idx).split("-")[0])
                        cursor.execute("INSERT INTO ProgramTur VALUES (%s, %s)", (p_id, t_id))
                
                conn.commit(); conn.close(); messagebox.showinfo("Başarılı", "Katalog verisi başarıyla güncellendi!"); fm.destroy(); self.raporla()
            except Exception as e: messagebox.showerror("Hata", str(e))
            
        tk.Button(fm, text="Değişiklikleri Kaydet", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 10, "bold"), bd=0, command=db_g).pack(fill="x", padx=30, ipady=6, pady=15)

    def tur_yonetim_penceresi(self):
        tw = tk.Toplevel(self); tw.geometry("340x350"); tw.configure(bg="#101726"); tw.title("Tür Yönetimi")
        tk.Label(tw, text="Yeni Tür Adı:", fg="white", bg="#101726").pack(pady=5)
        ent = tk.Entry(tw, width=20); ent.pack()
        lt = tk.Listbox(tw, bg="#1a233a", fg="white", bd=0); lt.pack(fill="both", expand=True, padx=20, pady=5)
        def l_t():
            lt.delete(0, tk.END)
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT tur_id, tur_adi FROM Tur")
            for r in cursor.fetchall(): lt.insert(tk.END, f"{r[0]}- {r[1]}")
            conn.close()
        def t_e():
            if not ent.get().strip(): return
            try:
                conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("INSERT INTO Tur (tur_adi) VALUES (%s)", (ent.get().strip(),))
                conn.commit(); conn.close(); l_t()
            except Exception as e: messagebox.showerror("Hata", str(e))
        def t_s():
            if not lt.curselection(): return
            t_id = int(lt.get(lt.curselection()[0]).split("-")[0])
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT COUNT(*) FROM ProgramTur WHERE tur_id=%s", (t_id,))
            if cursor.fetchone()[0] > 0: return messagebox.showerror("Hata", "Bu türe bağlı içerikler var!"), conn.close()
            cursor.execute("DELETE FROM Tur WHERE tur_id=%s", (t_id,)); conn.commit(); conn.close(); l_t()
        tk.Button(tw, text="Ekle", bg="#ff9f1c", command=t_e).pack(pady=2)
        tk.Button(tw, text="Sil", bg="#ff4d4d", fg="white", command=t_s).pack(pady=2); l_t()

    def kullanici_yonetim_penceresi(self):
        kw = tk.Toplevel(self); kw.geometry("450x380"); kw.configure(bg="#101726"); kw.title("Kullanıcı Yönetimi")
        tree_k = ttk.Treeview(kw, columns=("ID", "Ad Soyad", "Email", "Durum"), show="headings")
        for c in ("ID", "Ad Soyad", "Email", "Durum"): tree_k.heading(c, text=c); tree_k.column(c, width=110, anchor="center")
        tree_k.pack(fill="both", expand=True, padx=10, pady=10)
        def k_l(): #üyeleri listele
            for i in tree_k.get_children(): tree_k.delete(i)
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT kullanici_id, CONCAT(ad, ' ', soyad), email, durum FROM Kullanici WHERE rol_id=1")
            for r in cursor.fetchall(): tree_k.insert("", "end", values=r)
            conn.close()
        def t_d(): #üyenin durumu
            if not tree_k.selection(): return
            row = tree_k.item(tree_k.selection()[0], "values")
            new_d = "Pasif" if row[3] == "Aktif" else "Aktif"
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("UPDATE Kullanici SET durum=%s WHERE kullanici_id=%s", (new_d, row[0]))
            conn.commit(); conn.close(); k_l()
        
        def k_detay_goster(event):
            if not tree_k.selection(): return
            sel_k_id = tree_k.item(tree_k.selection()[0], "values")[0]
            detay_w = tk.Toplevel(kw); detay_w.geometry("520x450"); detay_w.configure(bg="#101726"); detay_w.title("Kullanıcı Detay Kartı")
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT ad, soyad, email, DATE_FORMAT(dogum_tarihi, '%Y-%m-%d'), ulke, durum FROM Kullanici WHERE kullanici_id=%s", (sel_k_id,))
            kd = cursor.fetchone()
            cursor.execute("SELECT SUM(izleme_suresi) FROM IzlemeLog WHERE kullanici_id=%s", (sel_k_id,))
            k_sure = cursor.fetchone()[0] or 0
            tk.Label(detay_w, text=f"{kd[0]} {kd[1]} - Hesap Bilgileri", font=("Arial", 12, "bold"), fg="#ff9f1c", bg="#101726").pack(pady=10)
            tk.Label(detay_w, text=f"E-mail: {kd[2]}  |  Ülke: {kd[4]}  |  Durum: {kd[5]}", fg="white", bg="#101726").pack()
            tk.Label(detay_w, text=f"Doğum Tarihi: {kd[3]}  |  Toplam İzleme Hacmi: {k_sure} Dakika", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).pack(pady=5)
            tk.Label(detay_w, text="Kullanıcının Detaylı İzleme Günlüğü (Geçmiş):", font=("Arial", 10, "bold"), fg="white", bg="#101726").pack(pady=(10,2))
            tree_kh = ttk.Treeview(detay_w, columns=("Yapım", "Tarih", "Bölüm", "Süre", "Durum"), show="headings", height=6)
            for ch in ("Yapım", "Tarih", "Bölüm", "Süre", "Durum"): tree_kh.heading(ch, text=ch); tree_kh.column(ch, width=95, anchor="center")
            tree_kh.pack(fill="both", expand=True, padx=10, pady=5)
            cursor.execute("SELECT p.program_adi, il.izleme_tarihi, il.izlenen_bolum, il.izleme_suresi, CASE WHEN il.tamamlandi_mi = 1 THEN 'Bitti' ELSE 'Kaldı' END FROM IzlemeLog il JOIN Program p ON il.program_id = p.program_id WHERE il.kullanici_id=%s ORDER BY il.izleme_id DESC", (sel_k_id,))
            for rh in cursor.fetchall(): tree_kh.insert("", "end", values=rh)
            conn.close()

        tree_k.bind("<Double-1>", k_detay_goster)
        tk.Button(kw, text="Durum Değiştir (Aktif/Pasif)", bg="#ff9f1c", command=t_d).pack(pady=10); k_l()

    def yeni_icerik_ekle(self):
        fm = tk.Toplevel(self); fm.geometry("340x480"); fm.configure(bg="#101726"); fm.title("Yeni Yapım Ekle")
        
        tk.Label(fm, text="Yapım Adı:", fg="white", bg="#101726").pack()
        ent_ad = tk.Entry(fm, width=28, bg="#1a233a", fg="white", bd=0); ent_ad.pack(ipady=3, pady=2)
        
        tk.Label(fm, text="Açıklama:", fg="white", bg="#101726").pack()
        ent_ozet = tk.Entry(fm, width=28, bg="#1a233a", fg="white", bd=0); ent_ozet.pack(ipady=3, pady=2)
        
        tk.Label(fm, text="Kategori Tipi:", fg="white", bg="#101726").pack()
        cmb_p_tip = ttk.Combobox(fm, values=["Film", "Dizi"], width=25)
        cmb_p_tip.set("Film"); cmb_p_tip.pack(pady=2)
        
        tk.Label(fm, text="Bölüm Sayısı:", fg="white", bg="#101726").pack()
        ent_b_sayi = tk.Entry(fm, width=28, bg="#1a233a", fg="white", bd=0)
        ent_b_sayi.insert(0, "1"); ent_b_sayi.config(state="disabled"); ent_b_sayi.pack(ipady=3, pady=2)
        
        def tip_degisti(event):
            if cmb_p_tip.get() == "Film":
                ent_b_sayi.config(state="normal"); ent_b_sayi.delete(0, tk.END); ent_b_sayi.insert(0, "1"); ent_b_sayi.config(state="disabled")
            else:
                ent_b_sayi.config(state="normal")
        cmb_p_tip.bind("<<ComboboxSelected>>", tip_degisti)

        tk.Label(fm, text="Uzunluk (Dk):", fg="white", bg="#101726").pack()
        ent_uzunluk = tk.Entry(fm, width=28, bg="#1a233a", fg="white", bd=0); ent_uzunluk.pack(ipady=3, pady=2)
        
        tk.Label(fm, text="Yayın Yılı:", fg="white", bg="#101726").pack()
        ent_yil = tk.Entry(fm, width=28, bg="#1a233a", fg="white", bd=0); ent_yil.pack(ipady=3, pady=2)

        tk.Label(fm, text="Yapım Türleri (Çoklu Seçim İçin Ctrl):", fg="#00d2ff", bg="#101726", font=("Arial", 9, "bold")).pack(pady=(5,0))
        lst_t = tk.Listbox(fm, selectmode="multiple", height=3, width=28, bg="#1a233a", fg="white", bd=0, selectbackground="#00d2ff")
        lst_t.pack(pady=2)
        try:
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("SELECT tur_id, tur_adi FROM Tur")
            for r in cursor.fetchall(): lst_t.insert(tk.END, f"{r[0]}- {r[1]}")
            conn.close()
        except: pass

        def db_k():
            try:
                conn = get_db_connection(); cursor = conn.cursor()
                b_sayisi_val = 1 if cmb_p_tip.get() == "Film" else int(ent_b_sayi.get())
                cursor.execute("INSERT INTO Program (program_adi, aciklama, program_tipi, bolum_sayisi, program_uzunlugu, yayin_yili) VALUES (%s,%s,%s,%s,%s,%s)", 
                               (ent_ad.get(), ent_ozet.get(), cmb_p_tip.get(), b_sayisi_val, int(ent_uzunluk.get()), int(ent_yil.get())))
                new_prog_id = cursor.lastrowid
                
                sel_t = lst_t.curselection()
                for idx in sel_t:
                    t_id = int(lst_t.get(idx).split("-")[0])
                    cursor.execute("INSERT INTO ProgramTur VALUES (%s, %s)", (new_prog_id, t_id))
                    
                conn.commit(); conn.close(); messagebox.showinfo("Başarılı", "Eklendi!"); fm.destroy(); self.raporla()
            except Exception as e: messagebox.showwarning("Hata", str(e))
            
        tk.Button(fm, text="Kataloğa Gönder", bg="#ff9f1c", fg="#0a0f1d", font=("Arial", 11, "bold"), bd=0, command=db_k).pack(pady=10, ipady=5, fill="x", padx=20)

    def secili_yapimi_sil(self):
        if not self.tree_admin.selection(): return messagebox.showwarning("Uyarı", "Katalogdan bir yapım seçin!")
        sel = self.tree_admin.selection()[0]
        s_id, y_adi = int(self.tree_admin.item(sel, "values")[0]), self.tree_admin.item(sel, "values")[1]
        if not messagebox.askyesno("Onay", f"'{y_adi}' silinsin mi?"): return
        try:
            conn = get_db_connection(); cursor = conn.cursor(); cursor.execute("DELETE FROM Program WHERE program_id = %s", (s_id,))
            conn.commit(); conn.close(); messagebox.showinfo("Başarılı", "Silindi."); self.raporla()
        except Error as e: messagebox.showerror("Hata", str(e))

    def raporla(self):
        for i in self.tree_admin.get_children(): self.tree_admin.delete(i)
        self.txt_panel.delete("1.0", tk.END)
        try:
            conn = get_db_connection(); cursor = conn.cursor()
            cursor.execute("SELECT program_id, program_adi, program_tipi, yayin_yili, bolum_sayisi FROM Program")
            for p in cursor.fetchall(): self.tree_admin.insert("", "end", values=p)
            
            sorgular = {
                "u_sayisi": "SELECT COUNT(*) FROM Kullanici WHERE rol_id=1",
                "i_sayisi": "SELECT COUNT(*) FROM IzlemeLog",
                "top_izlenen": "SELECT p.program_adi, COUNT(il.program_id) FROM IzlemeLog il JOIN Program p ON il.program_id=p.program_id GROUP BY il.program_id ORDER BY 2 DESC LIMIT 10",
                "top_puanli": "SELECT p.program_adi, ROUND(AVG(kp.puan),1) FROM KullaniciProgram kp JOIN Program p ON kp.program_id=p.program_id WHERE kp.puan IS NOT NULL GROUP BY kp.program_id ORDER BY 2 DESC LIMIT 10",
                "turler": "SELECT t.tur_adi, COUNT(il.izleme_id) FROM IzlemeLog il JOIN ProgramTur pt ON il.program_id=pt.program_id JOIN Tur t ON pt.tur_id=t.tur_id GROUP BY t.tur_id ORDER BY 2 DESC LIMIT 5",
                "aktifler": "SELECT CONCAT(u.ad, ' ', u.soyad), SUM(il.izleme_suresi) FROM IzlemeLog il JOIN Kullanici u ON il.kullanici_id=u.kullanici_id GROUP BY il.kullanici_id ORDER BY 2 DESC LIMIT 5",
                "son7": "SELECT COUNT(*) FROM IzlemeLog WHERE izleme_tarihi >= DATE_SUB(CURDATE(), INTERVAL 7 DAY)",
                "p_sayisi": "SELECT COUNT(*) FROM KullaniciProgram WHERE puan IS NOT NULL"
            }
            res = {}
            for k, sql in sorgular.items():
                cursor.execute(sql)
                res[k] = cursor.fetchall()
            
            ck = "========================================================================================\n"
            ck += "                        YONETICI RAPORLARI                 \n"
            ck += "========================================================================================\n"
            ck += f" Rapor Zamanı: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Üye Sayısı: {res['u_sayisi'][0][0]} Aktif\n"
            ck += f" İzlenme Sayısı: {res['i_sayisi'][0][0]} Kez | Değerlendirme Puanı Sayısı: {res['p_sayisi'][0][0]}\n"
            ck += f" Son 7 Gün İzleme Trafiği: {res['son7'][0][0]} İzleme\n"
            ck += "----------------------------------------------------------------------------------------\n"
            ck += "\nEN ÇOK İZLENEN 10 İÇERİK:\n"
            for idx, r in enumerate(res['top_izlenen'], 1): ck += f"  {idx}. {r[0]:<30} -> {r[1]} İzlenme\n"
            ck += "\nEN YÜKSEK PUAN ALAN 10 İÇERİK:\n"
            for idx, r in enumerate(res['top_puanli'], 1): ck += f"  {idx}. {r[0]:<30} -> Puan: {r[1]}/10\n"
            ck += "\nEN ÇOK İZLENEN TÜRLER:\n"
            for r in res['turler']: ck += f"  • {r[0]:<25} -> {r[1]} Log\n"
            ck += "\nEN AKTİF SEYİRCİLER:\n"
            for r in res['aktifler']: ck += f"  • {r[0]:<25} -> {r[1]} Dakika\n"
            self.txt_panel.insert(tk.END, ck); conn.close()
        except Error as e: self.txt_panel.insert(tk.END, f"Raporlama Hatası: {e}")

if __name__ == "__main__":
    app = NetflixApp()
    app.mainloop()