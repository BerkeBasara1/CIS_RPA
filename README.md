# CIS_RPA Calistirma Rehberi

Bu proje 3 ana script ile ilerleyen bir akistir:

1. `cis_rpa.py`
2. `indir_excel_url_gorseller_embed.py`
3. `gorsel_kopyalama_pipeline_all.py`

Dogru sira genelde **1 -> 2 -> 3** seklindedir.

Bu README'de:

- Hangi script ne zaman calistirilir
- Her adimdan once ne kontrol edilmelidir
- Hata alindiginda ne yapilmalidir
- Hangi durumda tum sureci bastan almak gerekir, hangi durumda sadece ilgili adimi yeniden calistirmak yeterlidir

anlatilmistir.

---

## 1. Projenin Kisa Ozeti

Bu proje temelde su isi yapar:

1. `cis_rpa.py` ile Skoda panelinden her sasi icin gorsel URL'leri toplanir.
2. `indir_excel_url_gorseller_embed.py` ile bu URL'ler gercek resim dosyalarina indirilir.
3. `gorsel_kopyalama_pipeline_all.py` ile unique araclardan indirilen gorseller, `Eşleşme ID` bazinda tum pipeline araclarina dagitilir.

Kisaca:

- Ilk script URL toplar
- Ikinci script dosya indirir
- Ucuncu script kopyalama / dagitim yapar

---

## 2. Onerilen Calistirma Sirasi

### Adim 1: `cis_rpa.py`

Amaci:

- Acik Excel dosyasindaki sasileri tek tek alip
- Skoda admin panelinde preview olusturmak
- Dis ve ic goruntulere ait URL'leri Excel'e yazmak

### Adim 2: `indir_excel_url_gorseller_embed.py`

Amaci:

- Excel'e yazilmis URL'leri okumak
- Her sasi icin klasor olusturmak
- Resimleri bu klasorlere indirmek

### Adim 3: `gorsel_kopyalama_pipeline_all.py`

Amaci:

- `Unique` sheet'teki kaynak gorselleri bulmak
- `Pipeline All` sheet'teki ilgili araclara kopyalamak
- Dagitimi `Eşleşme ID` uzerinden yapmak

---

## 3. Baslamadan Once Kontrol Listesi

Tum surece baslamadan once sunlari kontrol edin:

### Ortam

- Windows makine kullaniyor olun.
- Microsoft Excel masaustu surumu kurulu olsun.
- Google Chrome kurulu olsun.
- Python ortaminda gerekli kutuphaneler yuklu olsun.

### Olası paketler

Asagidaki paketler gerekli olabilir:

```bash
pip install selenium webdriver-manager pywin32 pyperclip pyautogui pandas openpyxl requests
```

### Dosya ve yol kontrolleri

Scriptlerde bircok yol kod icine sabit yazilmis durumda. Calistirmadan once mutlaka kontrol edin:

- `cis_rpa.py`
  - Skoda kullanici adi / sifre
  - `excel_gorsel.png` yolu
- `indir_excel_url_gorseller_embed.py`
  - `EXCEL_PATH`
  - `OUTPUT_DIR`
- `gorsel_kopyalama_pipeline_all.py`
  - `EXCEL_PATH`
  - `IMAGES_ROOT`
  - `OUT_ROOT`

### Cok onemli uyumluluk notu

Su an kodlarda varsayilan klasorler birebir ayni degil:

- `indir_excel_url_gorseller_embed.py` ciktisini `indirilen_gorseller3` altina aliyor.
- `gorsel_kopyalama_pipeline_all.py` ise varsayilan olarak `indirilen_gorseller2` altini okuyor.

Yani 2. adimdan sonra 3. adima gececekseniz, `gorsel_kopyalama_pipeline_all.py` icindeki `IMAGES_ROOT` degerini gercek kullandiginiz indirme klasorune gore guncellemeniz gerekir.

---

## 4. Script Bazli Detayli Calisma Rehberi

## 4.1 `cis_rpa.py`

### Bu script ne yapar?

Bu script:

- Chrome aciyor
- Skoda admin paneline login oluyor
- Acik Excel dosyasinin aktif sayfasindaki `E` sutunundaki sasileri okuyor
- Her sasi icin panelden gorsel URL'lerini topluyor
- URL'leri ayni Excel satirina yaziyor

Yazdigi kolonlar:

- `Araç Dış Arka Görsel Yolu`
- `Araç Dış Ön Görsel Yolu`
- `Araç Dış Yan Görsel Yolu`
- `Araç İç Arka Görsel Yolu`
- `Araç İç Ön Görsel Yolu`
- `Araç İç Yan Görsel Yolu`

### Calistirmadan once

- Excel dosyasi acik olsun.
- Dogru workbook aktif olsun.
- Dogru sheet aktif olsun.
- Ilk islenecek sasi `E2` hucresinde bulunsun.
- Excel ikon gorseli (`excel_gorsel.png`) ekranda gorulebilir olmasa bile dosya yolu dogru olsun.
- Skoda paneline erisim icin VPN ya da kurumsal ag gerekiyorsa aktif olsun.

### Nasil calistirilir?

```bash
python cis_rpa.py
```

### Beklenen sonuc

Script bittiginde:

- Excel'deki her sasi satirina 6 adet gorsel URL'si yazilmis olur
- Bir sonraki sasiye otomatik gecilir
- Bos satira gelince dongu durur

### Bu script calisirken dikkat edilmesi gerekenler

- Mouse'u ekran kosesine goturmeyin. `pyautogui` fail-safe tetiklenebilir.
- Excel penceresini kapatmayin.
- Aktif workbook'u degistirmeyin.
- Tarayiciyi elle kapatmayin.
- Ekran olcegi veya tema degisikse ikon eslestirme sapabilir.

### Sık karsilasilan hatalar ve cozumleri

#### 1. Login ekrani bulunamiyor

Belirti:

- Username / password alanlari bulunamiyor
- Timeout hatasi geliyor

Ne yapilmali:

- Kurumsal ağa / VPN'e bagli oldugunuzdan emin olun
- URL'nin erisilebilir oldugunu tarayicida manuel kontrol edin
- Login sayfasinin HTML'i degismis olabilir; selector'larin guncellenmesi gerekebilir
- Site yavas aciliyorsa bekleme suresini arttirin

#### 2. Excel COM nesnesine baglanamiyor

Belirti:

- Excel application bulunamiyor
- Aktif workbook yok

Ne yapilmali:

- Excel'i manuel acin
- Hedef dosyanin acik oldugundan emin olun
- Dogru sheet'i aktif hale getirin
- Tek bir workbook ile test edin

#### 3. Excel ikonu bulunamiyor

Belirti:

- `Excel ikonu belirtilen sure icinde bulunamadi`

Ne yapilmali:

- `excel_gorsel.png` dosya yolunu kontrol edin
- Ikon gorselini guncel ekran goruntusune gore yeniden alin
- Ekran olcegini 100% - 125% gibi daha stabil bir degere alin
- Excel pencerenizi taskbar veya masaustu durumuna gore test edin

#### 4. `PyAutoGUI FAIL-SAFE` hatasi

Belirti:

- Mouse ekran kosesine gidince script durur

Ne yapilmali:

- Mouse'u koseden cekin
- Scripti tekrar calistirin
- Mumkunse calisma sirasinda mouse hareketini minimumda tutun

#### 5. `Failed to get car preview`

Belirti:

- Script preview alamadigini soyluyor

Ne yapiliyor:

- Kod zaten bu durumda bir alt satirdaki sasiyi alip devam etmeye calisiyor

Yine de kontrol edilmesi gerekenler:

- Ilgili sasinin gercekten sistemde preview uretemedigi durum olabilir
- Bu satiri sonradan Excel'de filtreleyip manuel kontrol edin
- URL kolonlari bos kalan satirlari ayri raporlayin

#### 6. Slider'da 3 gorsel bulunamiyor

Belirti:

- `Dış tasarım slider’da 3 görsel yok`
- `İç tasarım slider’da 3 görsel yok`

Ne yapilmali:

- Site arayuzu degismis olabilir
- CSS selector'larin guncellenmesi gerekebilir
- Ilgili arac icin render eksik olabilir
- Bu sasi daha sonra tekrar denenebilir

### Hata durumunda tekrar calistirma kurali

- Sadece birkaç satir bos kaldiysa: Excel'de ilk eksik satiri `E2` konumuna gelecek sekilde ayarlayip tekrar calistirabilirsiniz.
- Tum akista login veya selector problemi varsa: once sorunu cozun, sonra scripti yeniden calistirin.
- Ayni satirin ustune tekrar yazmak genelde sorun olmaz; URL'ler yeniden yazilabilir.

---

## 4.2 `indir_excel_url_gorseller_embed.py`

### Bu script ne yapar?

Bu script:

- Excel'deki URL kolonlarini okur
- Her sasi icin klasor olusturur
- Gorselleri internetten indirir
- Indirme logu uretir

### Calistirmadan once

- `cis_rpa.py` tamamlanmis olsun
- Excel dosyasinda URL kolonlari dolu olsun
- `EXCEL_PATH` dogru dosyayi gostersin
- `OUTPUT_DIR` dogru cikti klasorunu gostersin

### Nasil calistirilir?

```bash
python indir_excel_url_gorseller_embed.py
```

### Beklenen sonuc

- Her sasi icin ayri bir klasor olusur
- URL'lerdeki gorseller klasore kaydedilir
- `indirilenler_log.csv` olusur

### Basari kontrolu nasil yapilir?

Sunlari kontrol edin:

- Klasor sayisi, URL'si dolu sasi sayisina yakin mi?
- Her sasi klasorunde beklenen sayida gorsel var mi?
- Log dosyasinda `hata` veya `gecersiz_url` satirlari var mi?

### Sık karsilasilan hatalar ve cozumleri

#### 1. Excel okunamiyor

Belirti:

- `Excel okunamadı`

Ne yapilmali:

- `EXCEL_PATH` dogru mu kontrol edin
- Dosya acik ve kilitli olsa bile genelde okunur, ama bozuk dosya varsa farkli kopya ile deneyin
- Dosya adinda tasima veya yeniden adlandirma olduysa scripti guncelleyin

#### 2. Sasi / URL kolonlari bulunamiyor

Belirti:

- `Şasi/VIN sütunu bulunamadı`
- Gorsel yolu kolonlari bulunamiyor

Ne yapilmali:

- Excel kolon adlarini birebir kontrol edin
- Fazladan bosluk veya karakter farki varsa duzeltin
- Gerekirse scriptteki `VIN_COL` ve kolon isimlerini guncelleyin

#### 3. HTTP 403 / 404 / timeout

Belirti:

- Log'ta `HTTP 403`, `HTTP 404` ya da `hata: ... timeout`

Ne yapilmali:

- URL eskiyip gecersiz hale gelmis olabilir
- Internet baglantinizi kontrol edin
- Gerekirse o satirlar icin `cis_rpa.py` ile URL'leri yeniden alin
- Sunucu cok yavas ise `WORKERS` sayisini azaltin
- SSL problemi varsa son care olarak `INSECURE_SSL = True` deneyebilirsiniz

#### 4. Dosyalar eksik iniyor

Belirti:

- Klasorde beklenenden az gorsel var

Ne yapilmali:

- Log dosyasinda ilgili sasiyi aratin
- URL kolonlarinin hepsinin dolu oldugunu kontrol edin
- Problemli sasiler icin `cis_rpa.py` adimina geri donun

#### 5. Yeniden calistirinca dosyalar atlanıyor

Varsayilan davranis:

- Dosya zaten varsa script o dosyayi atlayabilir

Ne yapilmali:

- Temiz tekrar istiyorsaniz ilgili sasi klasorunu silip scripti yeniden calistirin
- Tum klasoru bastan almak istiyorsaniz `OUTPUT_DIR` altini temizleyip tekrar calistirin

### Hata durumunda tekrar calistirma kurali

- Sadece birkac sasi eksikse: sadece eksik klasorleri silip scripti tekrar calistirin
- URL'ler bozuksa: once `cis_rpa.py` ile URL'leri yenileyin, sonra bu scripti tekrar calistirin
- Log'ta yaygin hata varsa: once nedenini bulun, sonra toplu tekrar calistirin

---

## 4.3 `gorsel_kopyalama_pipeline_all.py`

### Bu script ne yapar?

Bu script:

- `Pipeline All` ve `Unique` sheet'lerini okur
- `Eşleşme ID` bazinda eslesme kurar
- `Unique` tarafindaki kaynak sasi klasorunu bulur
- O klasordeki tum dosyalari ilgili pipeline sasilerine kopyalar
- Sonunda rapor uretir

### Calistirmadan once

- Indirilmis gorsel klasorleri hazir olsun
- `EXCEL_PATH` icindeki workbookta `Unique` ve `Pipeline All` sheet'leri bulunsun
- Her iki sheette de `Eşleşme ID` ve `Şasi` kolonlari olsun
- `IMAGES_ROOT`, gercek kullandiginiz gorsel klasorunu gostersin
- `OUT_ROOT`, yazilabilir bir cikti klasoru olsun

### Nasil calistirilir?

```bash
python gorsel_kopyalama_pipeline_all.py
```

### Beklenen sonuc

- `OUT_ROOT` altinda her pipeline sasisi icin klasor olusur
- Kaynak klasorde ne varsa bu klasorlere kopyalanir
- `copy_report.xlsx` olusur
- `missing_unique_chassis.txt` olusur
- `folders_not_in_unique.txt` olusur

### Sık karsilasilan hatalar ve cozumleri

#### 1. Excel bulunamiyor

Belirti:

- `Excel bulunamadı`

Ne yapilmali:

- `EXCEL_PATH` degerini kontrol edin
- Dosyanin gercek konumunu tekrar teyit edin

#### 2. Gorsel root bulunamiyor

Belirti:

- `Görsel root bulunamadı`

Ne yapilmali:

- `IMAGES_ROOT` yolunu kontrol edin
- 2. adimda olusan klasorun burasi oldugundan emin olun
- `indirilen_gorseller2` ve `indirilen_gorseller3` karisikligini kontrol edin

#### 3. Kolon bulunamiyor

Belirti:

- `Kolon bulunamadı`

Ne yapilmali:

- Sheet kolonlarini kontrol edin
- `Eşleşme ID` ve `Şasi` kolon isimleri degismisse scriptte aday listeleri guncellenmeli

#### 4. Match ID Unique sheet'te yok

Belirti:

- Raporlarda `NO_MATCH_ID_IN_UNIQUE_SHEET`

Ne yapilmali:

- Unique sheet ile Pipeline All sheet ayni veri setine mi ait kontrol edin
- Unique olusturma asamasinda eksik kayit olabilir
- Eslesme ID dagitimini yeniden hesaplamak gerekebilir

#### 5. Kaynak klasor bulunamiyor

Belirti:

- `MATCH_ID_FOUND_BUT_NO_SOURCE_FOLDER_FOR_ANY_UNIQUE_CHASSIS`

Ne yapilmali:

- Unique sheet'teki kaynak sasi klasorunun gercekten var olup olmadigini kontrol edin
- 2. adimda bu sasi icin gorsel inmis mi bakin
- Sasi klasor adlari ile Excel'deki sasi degerleri birebir eslesiyor mu kontrol edin

#### 6. Kaynak klasor bos

Belirti:

- `SOURCE_FOLDER_EXISTS_BUT_EMPTY`

Ne yapilmali:

- Indirme adiminda dosyalar eksik kalmis olabilir
- Ilgili sasiyi tekrar indirip bu scripti yeniden calistirin

### Hata durumunda tekrar calistirma kurali

- Sadece belli sasiler eksikse: ilgili cikti klasorlerini silip scripti tekrar calistirin
- Match ID eslesmesi bozuksa: once veri tarafini duzeltin, sonra scripti yeniden calistirin
- Tum dagitimi bastan yapmak istiyorsaniz `OUT_ROOT` klasorunu temizleyip tekrar calistirin

---

## 5. Ucten Uca Onerilen Operasyon Sirasi

Gunluk operasyon icin pratik sira:

1. Gerekli Excel dosyalarinin guncel oldugunu kontrol edin.
2. `cis_rpa.py` icindeki login bilgileri ve `excel_gorsel.png` yolunu kontrol edin.
3. Excel'i acin, hedef sheet'i aktif yapin, ilk sasi `E2`'de olsun.
4. `cis_rpa.py` calistirin.
5. URL kolonlarinin doldugunu kontrol edin.
6. `indir_excel_url_gorseller_embed.py` icindeki `EXCEL_PATH` ve `OUTPUT_DIR` degerlerini kontrol edin.
7. `indir_excel_url_gorseller_embed.py` calistirin.
8. `indirilenler_log.csv` dosyasini kontrol edin.
9. `gorsel_kopyalama_pipeline_all.py` icindeki `EXCEL_PATH`, `IMAGES_ROOT`, `OUT_ROOT` degerlerini kontrol edin.
10. Gerekirse `IMAGES_ROOT` degerini 2. adimin gercek cikti klasorune cevirin.
11. `gorsel_kopyalama_pipeline_all.py` calistirin.
12. `copy_report.xlsx` ve txt raporlarini kontrol edin.

---

## 6. Hata Oldugunda Hangi Adima Geri Donulmeli?

### Senaryo 1: URL'ler eksik veya bos

Geri donulecek adim:

- `cis_rpa.py`

Sebep:

- Problem URL uretme asamasinda

### Senaryo 2: URL var ama resim klasoru eksik

Geri donulecek adim:

- `indir_excel_url_gorseller_embed.py`

Sebep:

- Problem indirme asamasinda

### Senaryo 3: Resimler var ama pipeline dagitimi eksik

Geri donulecek adim:

- `gorsel_kopyalama_pipeline_all.py`

Sebep:

- Problem kopyalama / eslesme asamasinda

### Senaryo 4: Unique ve Pipeline eslesmeleri bozuk

Geri donulecek adim:

- Veri hazirlama asamasina
- Gerekirse `Eşleşme ID` uretimini yeniden gozden gecirmeye

Sebep:

- Kod degil, veri modeli bozuk olabilir

---

## 7. Temiz Yeniden Calistirma Stratejisi

Eger surecin bir bolumu bozulursa su sekilde ilerlemek daha guvenlidir:

### Sadece `cis_rpa.py` sorunluysa

- Excel dosyasini duzeltin
- Eksik kalan satirdan devam edin
- Indirme ve kopyalamaya hemen gecmeyin, once URL'leri kontrol edin

### Sadece indirme sorunluysa

- Problemli sasi klasorlerini silin
- `indir_excel_url_gorseller_embed.py` scriptini tekrar calistirin
- Log dosyasini tekrar kontrol edin

### Sadece kopyalama sorunluysa

- `OUT_ROOT` altindaki problemli klasorleri silin
- `copy_report.xlsx` raporunu baz alip tekrar calistirin

### Her sey karistiysa

En guvenli yol:

1. URL'leri kontrol et
2. Indirme klasorunu temizle
3. Indirmeyi yeniden al
4. Dagitimi bastan yap

---

## 8. Bilinen Riskler

Bu projede su riskler vardir:

- Login selector'lari site degisirse kirilabilir
- Excel ikon eslestirmesi ekran duzenine baglidir
- Sabit dosya yollari makineye ozeldir
- `gorsel_kopyalama_pipeline_all.py` ile `indir_excel_url_gorseller_embed.py` varsayilan klasorleri farkli olabilir
- Excel kolon isimleri degisirse scriptler hata verebilir

Bu nedenle her calistirmadan once:

- dosya yolu
- sheet adi
- kolon adi
- output klasoru

kontrol edilmelidir.

---

## 9. Hizli Ozet

Kisa cevap:

1. `cis_rpa.py`
2. `indir_excel_url_gorseller_embed.py`
3. `gorsel_kopyalama_pipeline_all.py`

Hata olursa:

- URL sorunuysa 1. adıma don
- Indirme sorunuysa 2. adıma don
- Dagitim sorunuysa 3. adıma don
- Veri eslesmesi sorunuysa Excel ve `Eşleşme ID` mantigini kontrol et

---

## 10. Ek Not

Projede ayrica `canli_unique_liste_guncelle.py` adli ek bir script bulunuyor. Bu script canli unique listeyi guncellemek ve yeni gelen listeye grup ID dagitmak icin kullanilabilir. Bu README'nin ana akisi ise kullanici istegi dogrultusunda sadece su uc scriptin operasyon sirasina odaklanmistir:

- `cis_rpa.py`
- `indir_excel_url_gorseller_embed.py`
- `gorsel_kopyalama_pipeline_all.py`
