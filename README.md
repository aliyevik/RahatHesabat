# Rahat Hesabat

«Rahat Yaşam» bina idarəetmə şirkəti üçün hesabat proqramı.
Xüsusi olaraq **İlahə Əşrəfqızı** üçün hazırlanmışdır. Təqdim edir: **FMC** (fmc.az).

## Bölmələr

1. **Binalar üzrə mənfəət / zərər.** Apsiyondan çıxarılan 4 Gəlir/Gider raporundan hesabat qurur:
   gəlirlər (ödəniş və hesablama əsaslı) və xərclər (ödəniş və hesablama əsaslı).
   Aylar üzrə «Hesablanma | Ödəmə» sütunları, ilk və son qalıqlar, nəticə və pul vəsaitinin qalığı göstərilir.
2. **Debitor / kreditor borclar.** Apsiyonun «Detaylı Borç Listesi» faylından hesabat qurur:
   hər kontragent və ay üzrə hesablanma, ödəmə və qalıq. «DEVREDEN BORÇLAR» sətirləri ilk saldoya düşür.

«Hesabatı hazırla» basılanda hesabat bütün ekranı tutan ayrıca səhifədə açılır («← Geri qayıt» — formaya).
İstənilən rəqəmə basanda kartoçka açılır: o rəqəmin hansı Apsiyon sənədlərindən yarandığı göstərilir
(«← Hesabata qayıt» və ya Esc). Ekranda iki görünüş var: yan-yana (forma) və alt-alta (12 ay ekrana sığır).

Hesabatı PDF və Excel-ə albom formatında çıxarmaq olur. Yarımçıq aylar hesabata daxil edilmir.

## EXE faylı

Hər dəfə `main` budağına yeni dəyişiklik göndəriləndə GitHub Actions `RahatHesabat.exe` faylını
özü yığır. Faylı yükləmək üçün: **Actions → son iş → Artifacts → RahatHesabat-exe**.
`v1.0.0` kimi teq qoyulanda fayl **Releases** bölməsinə də əlavə olunur.

## Proqramçı üçün

```
pip install -r requirements.txt
python main.py
```

---

# ƏDV Hesabatı (İntiqam Əliyev üçün — təqdim edir FMC)

Ayrıca proqram (`EDV_Hesabati.exe`, kodu `edv/` qovluğunda). e-taxes-dan alınan alış və satış
qaimələri, ƏDV depozit çıxarışı və Kapital Bank çıxarışı olduğu kimi yüklənir; şirkət VÖEN-i fayllardan
oxunur və bütün faylların eyni şirkətə aid olduğu yoxlanılır.

* **Məlumat bazası** — şirkət (VÖEN) üzrə il boyu: ilkin qalıqlar bir dəfə yazılır, sonra hər ay yeni
  fayllar əlavə olunur (təkrarlar atılır, statusu dəyişən qaimə yenilənir). İstənilən ay üçün hesabat
  bazanın başlanğıcından bütün hərəkətlərlə hesablanır — ayın ilk və son qalığı düzgün çıxır.
* **ƏDV yekunu** — aylar üzrə: alıcılardan banka daxil olan × 18% = hesablanmış ƏDV, depozitə daxil olan
  ƏDV, əvəzləşdirilən ƏDV (308), büdcəyə ödənilməli ƏDV, depozit qalığı.
* **Əvəzləşmə** — alış qaimələrinə ödənişlər VÖEN üzrə FIFO ilə bağlanır (bank → ƏDV-siz hissə,
  depozit → ƏDV). e-taxes «Əlavə 1» yükləmə faylı (308 | seriya | nömrə | ödənilmiş məbləğ).
* **Satış**, **Bank hesabatı**, **Kontragentlər** (üzləşmə aktı ilə), **Yoxlanılmalı** — aydın olmayan
  hər əməliyyat ayrıca qrupda.

### Baza, backup və kilid
* Baza interfeysdən ayrıdır: `%APPDATA%\EDVHesabati\baza\<VÖEN>.json`. Hər dəyişiklikdə avtomatik nüsxə
  `baza\backup\<VÖEN>\` qovluğuna və istifadəçinin seçdiyi əlavə qovluğa (D:\, flash, OneDrive) yazılır
  (son 40 nüsxə). «Nüsxədən bərpa et» ilə istənilən nüsxəyə qayıtmaq olur.
* İlkin qalıqlar (tarixi və məbləğləri) yadda saxlandıqdan sonra kilidlənir; dəyişmək üçün təsdiq kodu
  tələb olunur (kod proqramda yalnız heş şəklində saxlanılır). Bütün dəyişikliklər və uğursuz cəhdlər
  tarixçədə qalır.

### Hesabatlar (sifarişçinin Excel formasına uyğun)
* **ƏDV hesablaşması** — daxil olmuş məbləğ (pul nə vaxt gəlibsə həmin ay; ƏDV depozitə gələn ƏDV bankdakı
  puldan çoxdursa baza ƏDV÷18%), hesablanmış ƏDV, əvəzləşmə, büdcəyə ödənilməli/qaytarılan, ödənilmiş ƏDV,
  əvvəlki dövrdən qalan və növbəti dövrə keçən qalıq.
* **Təqdim edilən e-qaimə** — vergitutma ayı, Bank-1/ƏDV-1/Tarix, Bank-2/ƏDV-2/Tarix, debitor qalığı.
* **Gələn qaimə, Əvəzləşmə** — yol vergisi, azad/0%, ödənişlər, kreditor qalığı, əvəzləşmə dövrü, «Əlavə 1».
* **Bank əməliyyatları** — ay, növ, qarşı tərəf, kateqoriya filtrləri, axtarış, aid olduğu qaimə (tam/qismən).
* Hər hesabatda «Sütunlar» menyusu: lazımsız sütunları söndürmək (ekranda və PDF/Excel-də).

### v3.0 — modullar
* **Uçot modulları**: Göndərilən qaimələr, Gələn qaimələr, Bank, ƏDV depozit — fayllar öz bölməsində yüklənir;
  sətrə iki dəfə basanda redaktə (səbəb mütləqdir), «hesaba alınmasın», qaiməni əl ilə əlavə etmək.
  İdxal olunan ilkin məlumat dəyişmir — düzəlişlər ayrıca qatda saxlanılır, «Dəyişikliklər jurnalı»na yazılır.
* **ƏDV hesabatı** ayrıca pəncərədə (vərəqlər): Hesablaşma, Təqdim edilən e-qaimə, Gələn qaimə/Əvəzləşmə,
  Ödənişlər. Vergitutma ayı həm satışda, həm alışda ödənişin ayıdır.
* Dövr seçimi pəncərənin yuxarısındadır; dəyişəndə bütün bölmələr avtomatik yenilənir.
* **GitHub**: baza sıxılmış şəkildə gizli (Private) repoya avtomatik göndərilir (`baza/<VÖEN>.json.gz`),
  başqa kompüterdə «GitHub-dan yüklə». Açıq repoya göndərilmir.
