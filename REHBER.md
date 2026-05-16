# DIR Stok Takip ve Analiz Sistemi Kullanım Kılavuzu

Bu uygulama, Dahilde İşleme Rejimi (DIR) kapsamında ithal edilen hammaddelerin, ihracat beyannameleri ile FIFO (İlk Giren İlk Çıkar) prensibine göre eşleştirilmesini ve stok takibini sağlar.

## 🚀 Hızlı Başlangıç

1.  **İthalat Listesi Yükleme:** "Excel İşlemleri" menüsünden "İthalat Listesi Yükle" diyerek ithalat beyannamelerinizi sisteme aktarın.
2.  **İhracat Analizi:** Yine aynı menüden "İhracat Analizi Yükle" diyerek ihracat dosyanızı yükleyin. Sistem otomatik olarak hammadde düşümlerini yapacaktır.
3.  **Rapor Alma:** Analiz bittikten sonra "Excel Olarak İndir" butonu ile tüm eşleşmeleri ve hata raporlarını bilgisayarınıza kaydedebilirsiniz.

---

## 📊 Excel Format Gereksinimleri

Sistem, sütun başlıklarını esnek bir şekilde tanıyacak şekilde tasarlanmıştır.

### 1. İthalat Listesi (Girdiler)
Aşağıdaki kolonlardan en az biri bulunmalıdır:
*   **Beyanname No:** `TCGB Tescil No`, `Beyanname No` veya `Tescil No`
*   **Tarih:** `TCGB Tescil Tarihi`, `Beyanname Tarihi` veya `Tarih`
*   **Miktar:** `Miktar` (Sayı değeri)

### 2. İhracat Listesi (Çıktılar)
Sistem bu listedeki verileri işlerken stoktan düşüm yapar:
*   **Beyanname No:** `TCGB Tescil No`, `Beyanname No` veya `Tescil No`
*   **Tarih:** `TCGB Tescil Tarihi`, `Beyanname Tarihi` veya `Tarih`
*   **Ürün Kodu:** `Satır Kodu` veya `Kalem Kodu`
*   **Miktar:** `İhracat Miktarı` veya `Miktar`

---

## ⚙️ Teknik Kurallar ve Mantık

### FIFO (İlk Giren İlk Çıkar)
Sistem, bir ihracat işlemini karşılarken **en eski tarihli** ve stok bakiyesi olan ithalat beyannamesinden başlar. İhracatın hammadde ihtiyacı bir beyannameyi aşarsa, sıradaki beyannameye otomatik geçer.

### Tarih Kısıtlaması
Bir ihracat beyannamesi, ancak kendisinden **en az 1 gün önce** tescil edilmiş ithalat beyannameleri ile eşleşebilir. Aynı gün yapılan ithalatlar o ihracat için kullanılmaz.

### Sarfiyat Oranları (Verimlilik)
Sistemde "Satır Kodu" alanına göre aşağıdaki otomatik sarfiyat oranları uygulanır:
*   **26.1.01378.001:** 0.928
*   **26.1.01378.002:** 0.865

*Örnek: 100 AD ihracat yapıldığında, sistem (Miktar * Oran) formülü ile stoktan düşüm yapar.*

---

## 🔴 Parçalı Düşümler (Split Exports)

Sistem, bir ihracatın hammadde ihtiyacını mevcut en eski ithalattan karşılayamazsa otomatik olarak bir sonraki uygun ithalata geçer. Bu duruma **Parçalı Düşüm** denir.

*   **Arayüzde Görünüm:** Eşleşmeler sayfasında, birden fazla ithalattan beslenen ihraacat satırları **kırmızı arka plan** ile vurgulanır. Bu, o satırın tek bir beyanname ile kapatılamadığını belirtir.
*   **Excel Raporunda Görünüm:** Rapor dosyasında bu işlem için her bir ithalat bağlantısı ayrı bir satır olarak kaydedilir. Böylece hangi ithalattan ne kadar KG düşüldüğü net bir şekilde izlenebilir.

---

## ⚠️ Hata ve Uyarı Mesajları

*   **SORUNLU / HATA:** İhracatın gerçekleşebilmesi için yeterli hammadde stoğu bulunamamıştır veya tarihsel olarak uygun ithalat yoktur.
*   **KISMI:** İhracatın bir kısmı stoklardan karşılanmış ancak tamamı için stok yetmemiştir.
*   **BAŞARILI:** İhracat payı tam olarak stoktan düşülmüş ve ilgili ithalatlarla bağlanmıştır.

---

## 🧹 Veri Güvenliği
Uygulama tamamen tarayıcı tabanlı çalışır. "Tüm Verileri Sıfırla" butonu ile sistemi temizleyebilir ve yeni bir döneme başlayabilirsiniz.
