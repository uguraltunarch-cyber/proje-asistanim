# -*- coding: utf-8 -*-
"""
MarkItDown Dönüştürücü
PDF, Word, Excel, PowerPoint ve diğer dosyaları Markdown (.md) formatına çevirir.
Microsoft MarkItDown kütüphanesini kullanır.
"""
import os
import sys
import threading
import traceback
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# .exe olarak (pencereli modda) çalışırken konsol yoktur; bazı kütüphaneler
# ekrana yazmaya çalışınca çökmesin diye çıktıyı boşa yönlendiriyoruz.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

# Uygulamanın bulunduğu klasör (.exe ya da .py)
if getattr(sys, "frozen", False):
    UYGULAMA_KLASORU = os.path.dirname(sys.executable)
else:
    UYGULAMA_KLASORU = os.path.dirname(os.path.abspath(__file__))

# Sürükle-bırak desteği (varsa)
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_VAR = True
except Exception:
    DND_VAR = False

DESTEKLENEN = {
    ".pdf", ".docx", ".doc", ".pptx", ".xlsx", ".xls", ".csv", ".json", ".xml",
    ".html", ".htm", ".txt", ".epub", ".msg", ".zip",
    ".jpg", ".jpeg", ".png", ".wav", ".mp3",
}

# Renkler
ZEMIN = "#EEF1F4"
PANEL = "#FFFFFF"
METIN = "#1F2A36"
SOLUK = "#66737F"
VURGU = "#2F5D8A"
VURGU_KOYU = "#244A6E"
BASARI = "#2E7D4F"
HATA = "#B3261E"
YAZI = "Segoe UI"


class Uygulama:
    def __init__(self, kok):
        self.kok = kok
        self.kok.title("MarkItDown Dönüştürücü")
        ekran_y = self.kok.winfo_screenheight()
        ekran_g = self.kok.winfo_screenwidth()
        yuk = min(780, ekran_y - 90)
        gen = min(900, ekran_g - 60)
        self.kok.geometry(f"{gen}x{yuk}+{(ekran_g - gen) // 2}+{max(10, (ekran_y - yuk) // 2 - 30)}")
        self.kok.minsize(640, 560)
        self.kok.configure(bg=ZEMIN)

        self.dosyalar = []
        self.cikti_klasoru = tk.StringVar(value="")
        self.ayni_klasor = tk.BooleanVar(value=True)
        self.birlestir = tk.BooleanVar(value=False)
        self.calisiyor = False
        self.son_cikti = None

        self._stil()
        self._arayuz()

    # ---------- Görünüm ----------
    def _stil(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure("TFrame", background=ZEMIN)
        s.configure("Panel.TFrame", background=PANEL)
        s.configure("TLabel", background=ZEMIN, foreground=METIN, font=(YAZI, 10))
        s.configure("Panel.TLabel", background=PANEL, foreground=METIN, font=(YAZI, 10))
        s.configure("Baslik.TLabel", background=ZEMIN, foreground=METIN, font=(YAZI, 17, "bold"))
        s.configure("Alt.TLabel", background=ZEMIN, foreground=SOLUK, font=(YAZI, 10))
        s.configure("Bolum.TLabel", background=PANEL, foreground=METIN, font=(YAZI, 11, "bold"))
        s.configure("TButton", font=(YAZI, 10), padding=(12, 6))
        s.configure("Ana.TButton", font=(YAZI, 11, "bold"), padding=(20, 9),
                    background=VURGU, foreground="white", borderwidth=0)
        s.map("Ana.TButton",
              background=[("disabled", "#9AAABB"), ("active", VURGU_KOYU)],
              foreground=[("disabled", "white")])
        s.configure("TCheckbutton", background=PANEL, foreground=METIN, font=(YAZI, 10))
        s.map("TCheckbutton", background=[("active", PANEL)])
        s.configure("Horizontal.TProgressbar", troughcolor="#DDE3E9",
                    background=VURGU, borderwidth=0, thickness=8)

    def _arayuz(self):
        dis = ttk.Frame(self.kok, padding=20)
        dis.pack(fill="both", expand=True)

        ttk.Label(dis, text="Dosyaları Markdown'a çevir", style="Baslik.TLabel").pack(anchor="w")
        ttk.Label(dis, text="PDF, Word, Excel, PowerPoint, HTML ve daha fazlası. "
                            "Çıkan .md dosyaları yapay zeka araçlarına (AnythingLLM vb.) doğrudan verilebilir.",
                  style="Alt.TLabel", wraplength=760, justify="left").pack(anchor="w", pady=(2, 14))

        # --- Alt kısım: ilerleme + çalıştır ---
        alt = ttk.Frame(dis)
        alt.pack(side="bottom", fill="x", pady=(14, 0))
        self.ilerleme = ttk.Progressbar(alt, mode="determinate", style="Horizontal.TProgressbar")
        self.ilerleme.pack(fill="x")

        satir = ttk.Frame(alt)
        satir.pack(fill="x", pady=(8, 0))
        self.durum = ttk.Label(satir, text="Hazır", style="TLabel", foreground=SOLUK)
        self.durum.pack(side="left")
        self.calistir = ttk.Button(satir, text="Dönüştür", style="Ana.TButton", command=self._baslat)
        self.calistir.pack(side="right")
        self.ac_dugme = ttk.Button(satir, text="Klasörü aç", command=self._klasoru_ac, state="disabled")
        self.ac_dugme.pack(side="right", padx=8)

        # --- Ayarlar paneli ---
        p2 = ttk.Frame(dis, style="Panel.TFrame", padding=14)
        p2.pack(side="bottom", fill="x", pady=(12, 0))
        ttk.Label(p2, text="Kayıt yeri", style="Bolum.TLabel").grid(row=0, column=0, sticky="w", columnspan=3)

        ttk.Checkbutton(p2, text="Her dosyanın yanına kaydet (aynı klasöre)",
                        variable=self.ayni_klasor, command=self._ayar_degisti
                        ).grid(row=1, column=0, sticky="w", pady=(6, 2), columnspan=3)

        self.klasor_etiket = ttk.Label(p2, text="Klasör seçilmedi", style="Panel.TLabel", foreground=SOLUK)
        self.klasor_etiket.grid(row=2, column=0, sticky="w", columnspan=2)
        self.klasor_dugme = ttk.Button(p2, text="Kayıt klasörü seç", command=self._cikti_sec)
        self.klasor_dugme.grid(row=2, column=2, sticky="e")

        ttk.Checkbutton(p2, text="Hepsini tek bir .md dosyasında birleştir",
                        variable=self.birlestir, command=self._ayar_degisti
                        ).grid(row=3, column=0, sticky="w", pady=(8, 0), columnspan=3)
        p2.columnconfigure(1, weight=1)

        # --- Dosya listesi paneli ---
        p1 = ttk.Frame(dis, style="Panel.TFrame", padding=14)
        p1.pack(fill="both", expand=True)

        ust = ttk.Frame(p1, style="Panel.TFrame")
        ust.pack(fill="x")
        ttk.Label(ust, text="Dönüştürülecek dosyalar", style="Bolum.TLabel").pack(side="left")
        self.sayac = ttk.Label(ust, text="0 dosya", style="Panel.TLabel", foreground=SOLUK)
        self.sayac.pack(side="right")

        liste_cerceve = tk.Frame(p1, bg="#D5DCE3", bd=0)
        liste_cerceve.pack(fill="both", expand=True, pady=(8, 8))
        self.liste = tk.Listbox(liste_cerceve, font=(YAZI, 10), bg="#FAFBFC", fg=METIN,
                                selectbackground=VURGU, selectforeground="white",
                                activestyle="none", bd=0, highlightthickness=0,
                                selectmode="extended")
        kaydir = ttk.Scrollbar(liste_cerceve, orient="vertical", command=self.liste.yview)
        self.liste.configure(yscrollcommand=kaydir.set)
        self.liste.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        kaydir.pack(side="right", fill="y")

        self.bos_ipucu = tk.Label(self.liste, bg="#FAFBFC", fg=SOLUK, font=(YAZI, 10),
                                  text=("Dosyaları buraya sürükleyip bırakın\nya da aşağıdaki düğmeleri kullanın"
                                        if DND_VAR else "Aşağıdaki düğmelerle dosya veya klasör ekleyin"))
        self.bos_ipucu.place(relx=0.5, rely=0.5, anchor="center")

        if DND_VAR:
            self.liste.drop_target_register(DND_FILES)
            self.liste.dnd_bind("<<Drop>>", self._birakildi)

        dugmeler = ttk.Frame(p1, style="Panel.TFrame")
        dugmeler.pack(fill="x")
        ttk.Button(dugmeler, text="Dosya ekle", command=self._dosya_ekle).pack(side="left")
        ttk.Button(dugmeler, text="Klasör ekle", command=self._klasor_ekle).pack(side="left", padx=6)
        ttk.Button(dugmeler, text="Seçileni çıkar", command=self._secileni_cikar).pack(side="left")
        ttk.Button(dugmeler, text="Listeyi temizle", command=self._temizle).pack(side="right")

        self._ayar_degisti()

    # ---------- Liste işlemleri ----------
    def _ekle(self, yollar):
        eklenen, atlanan = 0, 0
        for y in yollar:
            y = os.path.normpath(y)
            if os.path.isdir(y):
                for kok, _, adlar in os.walk(y):
                    for ad in sorted(adlar):
                        tam = os.path.join(kok, ad)
                        if os.path.splitext(ad)[1].lower() in DESTEKLENEN and not ad.startswith("~$"):
                            if tam not in self.dosyalar:
                                self.dosyalar.append(tam); eklenen += 1
            elif os.path.isfile(y):
                if os.path.splitext(y)[1].lower() == ".md":
                    atlanan += 1
                elif y not in self.dosyalar:
                    self.dosyalar.append(y); eklenen += 1
        self._listeyi_yenile()
        if eklenen == 0 and yollar:
            self._durum("Desteklenen yeni dosya bulunamadı", SOLUK)
        else:
            self._durum(f"{eklenen} dosya eklendi", SOLUK)

    def _listeyi_yenile(self):
        self.liste.delete(0, "end")
        for d in self.dosyalar:
            self.liste.insert("end", f"  {os.path.basename(d)}     —  {os.path.dirname(d)}")
        n = len(self.dosyalar)
        self.sayac.configure(text=f"{n} dosya")
        if n:
            self.bos_ipucu.place_forget()
        else:
            self.bos_ipucu.place(relx=0.5, rely=0.5, anchor="center")

    def _birakildi(self, olay):
        self._ekle(self.kok.tk.splitlist(olay.data))

    def _dosya_ekle(self):
        turler = [("Desteklenen dosyalar", " ".join("*" + u for u in sorted(DESTEKLENEN))),
                  ("Tüm dosyalar", "*.*")]
        secim = filedialog.askopenfilenames(title="Dosya seç", filetypes=turler)
        if secim:
            self._ekle(secim)

    def _klasor_ekle(self):
        k = filedialog.askdirectory(title="Klasör seç (alt klasörler dahil taranır)")
        if k:
            self._ekle([k])

    def _secileni_cikar(self):
        for i in reversed(self.liste.curselection()):
            del self.dosyalar[i]
        self._listeyi_yenile()

    def _temizle(self):
        self.dosyalar.clear()
        self._listeyi_yenile()
        self.ilerleme["value"] = 0
        self._durum("Hazır", SOLUK)

    # ---------- Ayarlar ----------
    def _ayar_degisti(self):
        klasor_gerekli = (not self.ayni_klasor.get()) or self.birlestir.get()
        self.klasor_dugme.configure(state="normal" if klasor_gerekli else "disabled")
        if klasor_gerekli:
            k = self.cikti_klasoru.get()
            self.klasor_etiket.configure(text=k if k else "Klasör seçilmedi",
                                         foreground=METIN if k else SOLUK)
        else:
            self.klasor_etiket.configure(text="Klasör seçmeye gerek yok", foreground=SOLUK)

    def _cikti_sec(self):
        k = filedialog.askdirectory(title="Kayıt klasörü seç")
        if k:
            self.cikti_klasoru.set(os.path.normpath(k))
            self._ayar_degisti()

    # ---------- Dönüştürme ----------
    def _durum(self, metin, renk=METIN):
        self.durum.configure(text=metin, foreground=renk)

    def _baslat(self):
        if self.calisiyor:
            return
        if not self.dosyalar:
            messagebox.showinfo("Dosya yok", "Önce dönüştürülecek dosyaları ekleyin.")
            return
        if ((not self.ayni_klasor.get()) or self.birlestir.get()) and not self.cikti_klasoru.get():
            messagebox.showinfo("Kayıt klasörü", "Önce bir kayıt klasörü seçin.")
            return
        self.calisiyor = True
        self.calistir.configure(state="disabled")
        self.ac_dugme.configure(state="disabled")
        self.ilerleme.configure(maximum=len(self.dosyalar), value=0)
        threading.Thread(target=self._donustur, daemon=True).start()

    def _arayuzde(self, f, *a):
        self.kok.after(0, lambda: f(*a))

    def _donustur(self):
        self._arayuzde(self._durum, "MarkItDown yükleniyor…", SOLUK)
        try:
            from markitdown import MarkItDown
            md = MarkItDown(enable_plugins=False)
        except Exception as e:
            try:
                with open(os.path.join(UYGULAMA_KLASORU, "hata_kaydi.txt"), "w", encoding="utf-8") as f:
                    f.write(traceback.format_exc())
            except Exception:
                pass
            self._arayuzde(self._bitti, 0, [("MarkItDown", f"Kütüphane yüklenemedi: {e}")], None)
            return

        dosyalar = list(self.dosyalar)
        birlestir = self.birlestir.get()
        ayni = self.ayni_klasor.get()
        hedef_klasor = self.cikti_klasoru.get()
        basarili, hatalar, parcalar = 0, [], []

        for i, yol in enumerate(dosyalar, 1):
            ad = os.path.basename(yol)
            self._arayuzde(self._durum, f"Dönüştürülüyor ({i}/{len(dosyalar)}): {ad}", METIN)
            try:
                if hasattr(md, "convert_local"):
                    sonuc = md.convert_local(yol)
                else:
                    sonuc = md.convert(yol)
                metin = getattr(sonuc, "markdown", None) or getattr(sonuc, "text_content", "") or ""
                if birlestir:
                    parcalar.append(f"# {ad}\n\n{metin.strip()}\n")
                else:
                    klasor = os.path.dirname(yol) if ayni else hedef_klasor
                    cikti = os.path.join(klasor, os.path.splitext(ad)[0] + ".md")
                    with open(cikti, "w", encoding="utf-8") as f:
                        f.write(metin)
                    self.son_cikti = klasor
                basarili += 1
            except Exception as e:
                hatalar.append((ad, str(e) or e.__class__.__name__))
            self._arayuzde(self.ilerleme.configure, {"value": i})

        birlesik_yol = None
        if birlestir and parcalar:
            birlesik_yol = os.path.join(hedef_klasor, "birlesik_belgeler.md")
            try:
                with open(birlesik_yol, "w", encoding="utf-8") as f:
                    f.write("\n\n---\n\n".join(parcalar))
                self.son_cikti = hedef_klasor
            except Exception as e:
                hatalar.append(("birlesik_belgeler.md", f"Kaydedilemedi: {e}"))
                birlesik_yol = None

        self._arayuzde(self._bitti, basarili, hatalar, birlesik_yol)

    def _bitti(self, basarili, hatalar, birlesik_yol):
        self.calisiyor = False
        self.calistir.configure(state="normal")
        if self.son_cikti:
            self.ac_dugme.configure(state="normal")

        if hatalar and basarili == 0:
            self._durum("Hiçbir dosya dönüştürülemedi", HATA)
        elif hatalar:
            self._durum(f"{basarili} dosya dönüştürüldü, {len(hatalar)} dosyada hata var", HATA)
        else:
            ek = " (birlesik_belgeler.md)" if birlesik_yol else ""
            self._durum(f"{basarili} dosya dönüştürüldü{ek}", BASARI)

        if hatalar:
            ozet = "\n".join(f"• {ad}: {m[:180]}" for ad, m in hatalar[:12])
            if len(hatalar) > 12:
                ozet += f"\n… ve {len(hatalar) - 12} dosya daha"
            messagebox.showwarning("Dönüştürülemeyen dosyalar", ozet)

    def _klasoru_ac(self):
        if not self.son_cikti:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.son_cikti)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.son_cikti])
            else:
                subprocess.Popen(["xdg-open", self.son_cikti])
        except Exception as e:
            messagebox.showerror("Klasör açılamadı", str(e))


def main():
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)  # Windows'ta bulanık yazıyı önler
    except Exception:
        pass
    kok = TkinterDnD.Tk() if DND_VAR else tk.Tk()
    Uygulama(kok)
    kok.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        with open(os.path.join(UYGULAMA_KLASORU, "hata_kaydi.txt"),
                  "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        raise
