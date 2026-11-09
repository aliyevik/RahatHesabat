# Rahat Hesabat

«Rahat Yaşam» bina idarəetmə şirkəti üçün hesabat proqramı.
Xüsusi olaraq **İlahə Əşrəfqızı** üçün hazırlanmışdır. Təqdim edir: **FMC** (fmc.az).

## Bölmələr

1. **Binalar üzrə mənfəət / zərər.** Apsiyondan çıxarılan 4 Gəlir/Gider raporundan hesabat qurur:
   gəlirlər (ödəniş və hesablama əsaslı) və xərclər (ödəniş və hesablama əsaslı).
   Aylar üzrə «Hesablanma | Ödəmə» sütunları, ilk və son qalıqlar, nəticə və pul vəsaitinin qalığı göstərilir.
2. **Debitor / kreditor borclar.** Apsiyonun «Detaylı Borç Listesi» faylından hesabat qurur:
   hər kontragent və ay üzrə hesablanma, ödəmə və qalıq. «DEVREDEN BORÇLAR» sətirləri ilk saldoya düşür.

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
