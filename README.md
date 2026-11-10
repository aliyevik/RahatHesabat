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

# ƏDV Hesabatı (təqdim edir FMC)

Ayrıca proqram (`EDV_Hesabati.exe`, kodu `edv/` qovluğunda). e-taxes-dan alınan alış və satış
qaimələri, ƏDV depozit çıxarışı və Kapital Bank çıxarışı olduğu kimi yüklənir; şirkət VÖEN-i fayllardan
oxunur və bütün faylların eyni şirkətə aid olduğu yoxlanılır.

* **ƏDV yekunu** — aylar üzrə: alıcılardan banka daxil olan × 18% = hesablanmış ƏDV, depozitə daxil olan
  ƏDV, əvəzləşdirilən ƏDV (308), büdcəyə ödənilməli ƏDV, depozit qalığı.
* **Əvəzləşmə** — alış qaimələrinə ödənişlər VÖEN üzrə FIFO ilə bağlanır (bank → ƏDV-siz hissə,
  depozit → ƏDV). e-taxes «Əlavə 1» yükləmə faylı (308 | seriya | nömrə | ödənilmiş məbləğ).
* **Satış**, **Bank hesabatı**, **Kontragentlər** (üzləşmə aktı ilə), **Yoxlanılmalı** — aydın olmayan
  hər əməliyyat ayrıca qrupda.
