# PersonelTak-p

PersonelTak-p, üretim planlama ve lojistik ekiplerinin haftalık ve tespit bazlı çalışan skorlarını hesaplamasını sağlayan bir araçtır. Proje artık tamamen tarayıcı üzerinde çalışan bir HTML/JavaScript paneli içerir; böylece Python ortamı hazırlamadan aynı işlevleri sürdürebilirsiniz.

## Özellikler

- Yeni **tarayıcı tabanlı** PersonelTak paneli Python kurulumuna ihtiyaç duymadan çalışır.
- `Kriterler`, `Calisanlar` ve `Degerlendirmeler` sayfaları SheetJS ile okunarak önceki Python sürümündeki doğrulamalar korunur.
- Rol ve kategori ağırlıkları, tespit penceresi ve eksik kayıt eşiği aynı hesaplama kurallarıyla uygulanır.
- Haftalık kriterler ISO haftasına göre, tespit kriterleri seçilen tarih aralığına göre filtrelenir.
- Yeni değerlendirme kayıtları tarayıcı belleğinde tutulur ve tek tıkla güncel Excel dosyasını indirebilirsiniz.
- Statik dosya olduğu için paylaşımlı ağ klasörlerinden ya da intranet web sunucularından sorunsuz dağıtılabilir.

## Hızlı Başlangıç

1. Depoyu klonlayın veya ZIP olarak indirin.
2. Kök dizindeki `index.html` dosyasını çift tıklayarak açın; tarayıcı doğrudan PersonelTak panelini gösterecektir.
3. Alternatif olarak kök dizinde basit bir sunucu açabilirsiniz:
   ```bash
   python -m http.server 8000
   ```
   Ardından tarayıcıda `http://localhost:8000` adresine gidin. Varsayılan index dosyası paneli otomatik yükler.
4. Panelde ilk adım olarak PersonelTak Excel dosyanızı (`Kriterler`, `Calisanlar`, `Degerlendirmeler` sayfalarını içeren çalışma kitabı) yükleyin.
5. Rapor tarihini seçip **Raporu Hesapla** düğmesine basarak skor ve eksik listelerini görüntüleyin.
6. Yeni değerlendirme ekleyin ve **Excel İndir** düğmesiyle güncel dosyayı dışa aktarın.

> Uygulama SheetJS kütüphanesini CDN üzerinden yükler; ek kurulum gerekmez. Offline kullanım için `web/vendor` klasörüne yerel kopya koyup `index.html` içindeki `<script>` adresini güncelleyebilirsiniz.

### PowerShell ile hızlı kurulum

Ekstra HTML kütüphanesi yüklemeniz gerekmez; yalnızca Python 3 ile gelen yerleşik sunucu yeterlidir. Aşağıdaki komutları PowerShell penceresinde sırasıyla çalıştırabilirsiniz:

```powershell
Set-Location "C:\\calisma\\PersonelTak-p" # Depo klasörüne gir
python -m http.server 8000                    # Yerleşik statik sunucuyu başlat
Start-Process "http://localhost:8000"        # Paneli varsayılan tarayıcıda aç
```

Sunucuyu kapatmak için PowerShell penceresinde `Ctrl+C` kombinasyonunu kullanın.

## JavaScript Bağımlılıkları

| Kütüphane | İşlev |
| --- | --- |
| [SheetJS (`xlsx`)](https://sheetjs.com/) | Excel dosyalarını okuyup yazmak için kullanılır. |
| Yerleşik ES2020 API'leri | ISO hafta, tarih işlemleri ve tablo üretimi için kullanılır. |

## Panel Kullanımı

1. **Excel'i yükleyin:** Dosya yalnızca tarayıcı belleğinde tutulur; sunucuya gönderilmez.
2. **Rapor tarihini belirleyin:** Varsayılan olarak bugün seçilir. ISO hafta otomatik hesaplanır.
3. **Raporu hesaplayın:** Skor ve eksik tabloları güncellenir, olası tutarsızlıklar uyarı olarak listelenir.
4. **Yeni kayıt ekleyin:** Formu doldurup gönderdiğinizde kayıt veriye eklenir ve tablolar yenilenir.
5. **Excel'i indirin:** Güncel veri aynı çalışma kitabı yapısıyla indirilebilir.

## Eski Python Sürümleri

`src/` klasöründeki Python betikleri önceki sürümle geriye dönük uyumluluk sağlamak amacıyla korunmaktadır. Yeni geliştirmeler HTML/JS paneline odaklanmaktadır; ancak Python tabanlı CLI veya Flask sunucusuna ihtiyaç duyarsanız eski betikleri kullanabilirsiniz.
