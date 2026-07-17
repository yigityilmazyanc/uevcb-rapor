# UEVÇB Raporu — EPİAŞ Veri Arayüzü

EPİAŞ Şeffaflık Platformu'ndan organizasyon seçerek saatlik veri çeker ve
organizasyon adıyla TEK Excel dosyası üretir:

| Sütun | İçerik | EPİAŞ kaynağı |
|---|---|---|
| A, B | Tarih, Saat | — |
| C, D | GÖP Eşleşme Alış / Satış (MWh) | `dam-clearing` (org bazlı) |
| E, F | İA Alış / Satış (MWh) | `bi-long` / `bi-short` (org bazlı) |
| G, H | GİP Alış / Satış (MWh) | `idm-qty` (org bazlı, yalnız saatlik kontratlar) |
| I… | Her santralin İlk KGÜP / Son KGÜP çifti (MWh) | `kgup-v1` / `kgup` (**UEVÇB bazlı**) |
| … | UEVM Toplam (MWh) | `uevm` (santral bazlı; ~1,5 ay geriden yayınlanır) |
| … | EDM = UEVM − Satışlar + Alışlar | Excel formülü |
| son | Net Satış Miktarı = (GÖP+İA+GİP Satış) − (GÖP+İA+GİP Alış) | Excel formülü |

Not: EPİAŞ, GÖP eşleşme / İA / GİP verilerini yalnız **organizasyon** bazında
yayınlar; UEVÇB kırılımı yoktur (dosyanın 1. satırında da yazar). KGÜP ve UEVM
kolonları gerçekten santral bazlıdır. GİP'te blok kontratlar saate bölünemediği
için yalnız saatlik kontratlar sayılır.

Organizasyon listesi KGÜP (İlk Versiyon) bölümünden gelir ve seçilen tarih
aralığının **tamamı** taranır: ay içinde eklenen/çıkan/toplayıcı değiştiren
firmalar da listede görünür. UEVÇB listesi ay ay taranıp birleştirilir.

---

## 🔧 Kurulum (Windows) — 3 adım, tek sefer

### 1) Python kur
- https://www.python.org/downloads/ adresinden **Python 3.11+** indir.
- Kurulumda **"Add python.exe to PATH"** kutusunu MUTLAKA işaretle.

### 2) Bu repoyu indir
- Git varsa: `git clone https://github.com/<kullanici>/uevcb-rapor.git`
- Git yoksa: GitHub'da **Code → Download ZIP** → bir klasöre çıkart.

### 3) `kurulum.bat`'a çift tıkla  ← **gerisini bu yapar**
- Gerekli paketleri kurar (`eptr2 pandas openpyxl`),
- Masaüstüne **"UEVCB Raporu"** kısayolu koyar,
- `ayarlar.json`'u oluşturup Not Defteri'nde açar — iki alanı doldur, kaydet:

```json
{
  "api_kullanici": "EPOSTANIZI-BURAYA-YAZIN",
  "api_sifre": "SIFRENIZI-BURAYA-YAZIN"
}
```

- Bilgiler EPİAŞ Şeffaflık Platformu (kayit.epias.com.tr) hesabının e-postası ve
  şifresidir. Hesap yoksa oradan ücretsiz açılır.
- `ayarlar.json` **gitignore'dadır** — şifre GitHub'a asla çıkmaz. Yine de kimseyle paylaşma.
- İsteğe bağlı: `"api_sifre_dosyasi": "~/.epias_pw"` (şifreyi ayrı dosyadan okur;
  `EPIAS_PW` ortam değişkeni hepsini ezer).

## Kullanım
Masaüstündeki **"UEVCB Raporu"** kısayoluna çift tıkla (konsol penceresi açılmaz;
alternatif: `arayuz.bat`).

1. Açılışta tüm organizasyonlar otomatik listelenir — "Ara" kutusu yazdıkça süzer.
2. Organizasyona çift tıkla (ya da "UEVÇB'leri Göster") → UEVÇB adları sağda listelenir.
3. Üstten tarih aralığını ve kayıt klasörünü ayarla (varsayılan: Downloads).
4. **"Excel Oluştur"** → her UEVÇB için `<UEVÇB adı>.xlsx` yazılır.

Komut satırından da çalışır:
```bat
py uevcb_rapor.py --org MASAT --bas 2025-01-01 --bit 2026-07-17
```

Notlar:
- EPİAŞ hız limitine (429) takılırsa program bekleyip kendiliğinden yeniden dener.
- Yazılacak Excel dosyası açıksa o dosya atlanır ve uyarı verilir — kapatıp yeniden çalıştır.
