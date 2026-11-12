# Mühasibat sistemi — layihə planı

Sizin cavablarınız (A1–A7) və qalan suallar üzrə mənim tövsiyələrim əsasında.

## 1. Əsas qərarlar

| Mövzu | Qərar |
|---|---|
| Kimlər üçün | Sizin şirkətləriniz + sonra müştərilərə satış / xidmət (çox müştərili sistem) |
| İstifadəçilər | Bir ofisdə bir neçə nəfər + internetlə müxtəlif yerlərdən |
| Platforma | Brauzerdə işləyən sistem + Windows proqramı (eyni interfeys, pəncərədə açılır) |
| Baza | Serverdə PostgreSQL. Kiçik ofis üçün ofis kompüteri server ola bilər |
| Uçot standartı | Şirkət kartında seçilir: Milli standartlar (kommersiya), kiçik sahibkarlıq üçün sadələşdirilmiş, MHBS |
| Vergi rejimi | Şirkət kartında, tarixə görə: ƏDV + mənfəət, sadələşdirilmiş, fərdi sahibkar |
| Valyuta | Yalnız AZN (struktur valyutaya hazır qalır, sonra əlavə etmək olar) |
| Analitika | Filial, xərc mərkəzi, layihə / obyekt, anbar + kontragent, nomenklatura, işçi |
| Kontragentlər | 531 / 211 + kontragent analitikası (alt hesab açmadan) |
| Avanslar | Ayrıca hesablarda, qaimə gələndə avtomatik əvəzləşir |
| Maya dəyəri | FIFO (uçot siyasətində seçim: FIFO / orta çəkili) |
| Dövrün bağlanması | Ay bağlanır, açmaq yalnız admin kodu ilə, jurnalda qeyd |

## 2. Arxitektura

- **Server**: Python (FastAPI) + PostgreSQL. Bütün hesablamalar serverdə.
- **İnterfeys**: brauzerdə işləyən tətbiq. Windows proqramı onu ayrıca pəncərədə açır (ikinci interfeys yazılmır).
- **Çox müştərili**: hər müştəri (təşkilat) ayrıca; təşkilatın içində bir neçə şirkət.
- **Təhlükəsizlik**: istifadəçi rolları, hər modula icazə (baxış / yazma / təsdiq), bütün dəyişikliklərin jurnalı.
- **Ehtiyat nüsxə**: serverdə gündəlik avtomatik + istənilən vaxt Excel / JSON ixracı.

## 3. Hesablama məntiqi (bütün modullar üçün eyni)

Hər modul **sənəd** yaradır, sənəd təsdiqlənəndə **müxabirləşmələr** (jurnal yazılışları) avtomatik yaranır. Bütün hesabatlar yalnız jurnaldan qurulur.
Beləliklə borc, ƏDV, anbar qalığı, əmək haqqı borcu — hamısı eyni yerdən hesablanır və bir-birinə uyğun gəlir.

Sənəd → yazılış qaydaları cədvəldə saxlanılır və şirkət üzrə dəyişdirilə bilər.

## 4. Mərhələlər

| № | Mərhələ | Nə daxildir |
|---|---|---|
| 1 | Əsas | Server, baza, istifadəçilər və rollar, şirkətlər, hesablar planı (rəsmi + alt hesablar), analitika kataloqları, əl ilə müxabirləşmə, ilkin qalıqlar, dövriyyə cədvəli, hesab kartı, baş kitab, ayın bağlanması |
| 2 | Alış, satış, bank, kassa | E-qaimə yükləmə (alış / satış), bank çıxarışı (qaydaları öyrənən), ƏDV depozit, kassa orderləri, avanslar, 211 / 531 analitikası, ƏDV bəyannaməsi və Əlavə 1 (indiki ƏDV proqramının məntiqi jurnal üzərində), köhnə ƏDV bazasının köçürülməsi |
| 3 | Nomenklatura və anbar | Növlər, qruplar (ağac), vahidlər, barkod, partiya / son istifadə tarixi, qiymət növləri, çoxanbar, köçürmə, inventarizasiya, silinmə, FIFO maya dəyəri, alış xərclərinin paylanması, e-qaimə mallarının nomenklaturaya uyğunlaşdırılması, satış nomenklatura ilə və ya cəm |
| 4 | Əmək haqqı və kadr | İşçi kartı, ştat, tabel, hesablamalar (məzuniyyət, xəstəlik, mükafat, ezamiyyə), tutulmalar (dərəcələr tarixli cədvəldə), bank ödəniş faylı, DSMF / gəlir vergisi hesabatları, xərclərə paylanma |
| 5 | Əsas vəsaitlər | Kart, mühasibat və vergi amortizasiyası paralel, təmir, yenidən qiymətləndirmə, silinmə |
| 6 | İstehsalat | Texnoloji kart (versiyalı), yarımfabrikat, istehsal sifarişi, birbaşa və ümumi xərclər, planlı / faktiki maya dəyəri, brak və itkilər |
| 7 | Hesabatlar | Balans, mənfəət və zərər, pul vəsaitlərinin hərəkəti, kapitalda dəyişiklik, idarəetmə hesabatları (marja, büdcə — fakt, obyektlər üzrə mənfəət), Rahat Yaşam abunə hesablaması |

Hər mərhələ ayrıca təhvil verilir və işlək olur; növbəti mərhələ əvvəlkinin üzərində qurulur.
