# DIR Stok Takip Sistemi - Geliştirme Özeti

Bu belge, Dahilde İşleme Rejimi (DIR) kapsamında geliştirilen FIFO tabanlı Stok ve Sarfiyat Takip Sistemi'nin gelişim sürecini ve teknik özelliklerini özetlemektedir.

## 🚀 Proje Genel Bakışı
Dahilde İşleme Rejimi kapsamında ithal edilen ham maddelerin, FIFO (First-In, First-Out) prensibiyle ihracatlarla eşleştirilmesi, kalan stokların takibi ve detaylı raporlanması hedeflenmiştir.

---

## ✨ Ana Özellikler

### 1. Dinamik Veri Girişi
- **Excel İthalat (İthalat Listesi):** Ham madde girişlerini toplu olarak beyanname no, tarih ve miktar bazında sisteme aktarma.
- **Excel İhracat (İhracat Listesi):** Sarfiyat yapılacak ürün listesini toplu yükleme ve analiz için hazırda tutma.
- **Manuel İşlemler:** Tekil kayıtlar için kullanıcı dostu İthalat ve İhracat ekleme diyalogları.

### 2. FIFO Analiz Motoru
- İhracat beyanname tarihlerine göre kronolojik sıralama.
- Her bir ihracat satırı için uygun tarihli ithalat stoğu tespiti.
- **Parçalı Düşüm Desteği:** Bir ihracat kaleminin birden fazla ithalat beyannamesinden karşılanabilmesi (Split Export).
- Sarfiyat oranlarına (Consumption Rate) göre otomatik ham madde ihtiyacı hesaplama.

### 3. Kullanıcı Arayüzü (UI/UX)
- **Modern Dashboard:** Yan menü (Sidebar) destekli, Emerald ve Slate renk paletiyle kurumsal tasarım.
- **Akıllı Listeleme:** 
    - **Stok Sayfası:** Kalan miktarlar ve tüketim durumu.
    - **Eşleşmeler:** Hangi ihracatın hangi ithalatı tükettiğini gösteren detaylı görünüm.
    - **Hatalar:** Stok yetersizliği veya reçete sorunlarını gösteren hata tablosu.
- **Görsel İndikatörler:** Parçalı düşüm yapılan satırların belirginleşmesi için kırmızı vurgulu (Red-Alert) hücre tasarımı.

### 4. Raporlama Sistemi
- **Gelişmiş Excel Çıktısı:**
    - **Analiz Raporu Sayfası:** FIFO eşleşmelerinin satır satır dökümü.
    - **İthalat Listesi Sayfası:** Güncel stok durumu özeti.
    - **İhracat Listesi Sayfası:** Analize giren tüm ihracat kalemleri.

---

## 🛠 Teknik Geliştirmeler ve Son Düzeltmeler

### Yapılan Son İyileştirmeler:
- **Analiz Süreci Revizyonu:** "İhracat Listesi Yükle" ve "Analizi Çalıştır" butonları birbirinden ayrıldı. Dosya yüklendiğinde "Bekleyen Veri" olarak saklanıp, butonla analiz tetiklenir hale getirildi.
- **Listeleme Görünümü:** İlgili ithalatlar kısmında sadece beyanname numarası kalacak şekilde sadeleştirme yapıldı.
- **Parçalı İthalat Düşümü (Split Export):** 
    - Bir ihracat kaleminin ihtiyacı tek bir ithalatla karşılanamadığında, FIFO sırasına göre sonraki beyannamelere otomatik geçiş özelliği eklendi.
    - Bu "parçalı" işlemler UI üzerinde kırmızı arka plan ile vurgulandı.
    - Excel raporunda her parçalı düşüm ayrı satır olarak dökülerek şeffaflık sağlandı.
- **Excel Çıktı Yapısı:** Eksik olan Sheets (İthalat ve İhracat listeleri) ana rapor dosyasının içine eklendi.
- **Hata Giderme:** "FileText is not defined" gibi kütüphane/ikon eksiklikleri giderilerek sistem kararlı hale getirildi.

---

## 📋 Teknik Detaylar
- **Framework:** React 18 (Vite)
- **Styling:** Tailwind CSS (Modern UI)
- **Data Processing:** XLSX (SheetJS)
- **Animations:** Framer Motion
- **Icons:** Lucide-React
