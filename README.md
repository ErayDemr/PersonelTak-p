# PersonelTak-p

PersonelTak-p, üretim planlama ve lojistik ekipleri için haftalık ve tespit bazlı çalışan skorlarını hesaplayan bir Python aracıdır. Kaynak Excel dosyasını okuyarak rol ve kategori ağırlıklarına göre 0-100 aralığında skorlar üretir, eksik değerlendirmeleri raporlar ve çıktıyı Excel olarak dışa aktarır.

## Özellikler

- `Kriterler`, `Calisanlar` ve `Degerlendirmeler` sayfalarını okuyarak veri doğrulama yapar; çalışan listesi istenirse standart ağ dizininden okunur.
- Rol ve kategori ağırlıklarını, tespit penceresini ve rapor yollarını yapılandırma dosyasından okur.
- Haftalık kriterler için içinde bulunulan ISO haftasını, tespit kriterleri için konfigüre edilebilir gün penceresini dikkate alır.
- Eksik veya yetkisiz değerlendirmeleri dosya tabanlı log'a yazar; çok kullanıcı aynı anda çalışırken otomatik dosya kilitleme uygular.
- CLI komutu ile skorları hesaplar, `rapor_YYYY-Www.xlsx` dosyası ile birlikte isteğe bağlı CSV ve Power BI veri seti üretir.
- Tüm işlevsellik `src/personeltak_app.py` dosyasında tek bir kod tabanında toplanmıştır; böylece dağıtım ve `.exe` üretimi sadeleşir.
- Flask tabanlı HTML arayüz sayesinde raporları tarayıcıdan görüntüleyip yeni puan girişleri yapılabilir.

## Hızlı Başlangıç

Tek komutla kurulum ve rapor üretimi yapmak için `COMMAND.md` dosyasındaki PowerShell veya Bash komutunu kullanabilirsiniz. Komut, sanal ortam kurulumundan raporlamaya kadar tüm adımları tek satırda yürütür.

## Kurulum

```bash
python -m venv .venv
source .venv/bin/activate  # Windows için .venv\\Scripts\\activate
pip install -r requirements.txt  # alternatif: pip install pandas openpyxl pyyaml filelock flask
```

> Not: Depoda kilitli bağımlılık dosyası bulunmadığından gerekli paketleri manuel kurun.

## Yapılandırma

Varsayılan ayarlar `config.example.yaml` dosyasında gösterilmiştir.

- `role_weights`: Rol bazlı ağırlıklar.
- `category_weights`: Kategori bazlı ağırlıklar.
- `tespit_days`: Tespit kriterleri için gün penceresi.
- `timezone`: Raporlamada kullanılacak saat dilimi.
- `excel_path`: Varsayılan kaynak Excel dosyası.
- `employees_path`: Çalışan listesinin tutulduğu standart ağ/yol (örn. `C:/ProgramData/PersonelTak/Calisanlar.xlsx`).
- `report_path`: Raporların yazılacağı klasör.
- `log_path` & `log_level`: Döner log dosyası konumu ve seviyesi.
- `missing_threshold`: Eksik kriter sayısı eşik değeri (opsiyonel).
- `csv_export`: Excel'e ek olarak CSV çıktısı üretimi.
- `powerbi_export` ve `powerbi_output`: Power BI veri seti klasörü ve etkinleştirme bayrağı.
- `lock_timeout`: Çok kullanıcılı erişim için dosya kilidi bekleme süresi (saniye).

Özel bir yapılandırma dosyasını `--config` parametresi ile CLI'ya verebilirsiniz.

## Komut Satırı Kullanımı

```bash
python src/personeltak_app.py summarize --output raporlar/
```

Parametreler:

- `excel`: Kaynak Excel dosyasının yolu (boş bırakılırsa yapılandırmadaki `excel_path` kullanılır). Proje kurulumundan sonra `personeltak summarize` komutu aynı argümanlarla kullanılabilir.
- `--config`: YAML/JSON konfigürasyon dosyası.
- `--asof`: ISO tarih formatında rapor tarihi (örn. `2024-03-22`).
- `summarize --output`: Raporların yazılacağı klasör.
- `record`: Yeni bir değerlendirme satırı eklemek için kullanılır.
- `web`: Aynı yapılandırmayı kullanarak HTML dashboard'u çalıştırır.

Örnek kayıt ekleme:

```bash
python src/personeltak_app.py record --sicil 123 --rol Şef --po 1 --puan 4.5 --tarih 2024-03-18
```

Komutlar çalıştırıldığında `logs/personeltak.log` dosyasına eklenir ve aynı dosya üzerinde birden fazla kullanıcı çalıştığında otomatik kilitleme uygulanır.

## HTML Arayüzü (Flask)

Raporları tarayıcı üzerinden izlemek ve yeni değerlendirme girişi yapmak için yerleşik Flask sunucusunu kullanabilirsiniz.

```bash
python src/personeltak_app.py web --config config.yaml --host 0.0.0.0 --port 8050
```

- `--host` ve `--port` parametreleri kurumunuzdaki ağ üzerinden paylaşım yapabilmeniz için özelleştirilebilir.
- Web arayüzü, aynı Excel dosyasını birden fazla kullanıcı kullandığında bile dosya kilidi ve loglama politikalarını uygular.
- Sayfa üzerindeki “Excel İndir” düğmesi, ilgili haftanın Excel raporunu anlık olarak üretip indirmenizi sağlar.
- Formu kullanarak yeni değerlendirme satırlarını (sicil, rol, Po, puan, tarih/not) hızlıca ekleyebilirsiniz.

## CSV & Power BI Entegrasyonu

`csv_export: true` ayarı ile her rapor çalıştırıldığında skor ve eksik listeleri `rapor_YYYY-Www_Skorlar.csv` ve `rapor_YYYY-Www_EksikPuanlamalar.csv` olarak üretilir. `powerbi_export: true` ayarı etkinse, Power BI için `personeltak_powerbi_YYYY-Www.csv` dosyası güncellenir ve eksik sayısı bilgisi dahil edilir.

## Windows `.exe` Paketleme

Aracı uç kullanıcılar için tek dosya `.exe` olacak şekilde paketlemek amacıyla [PyInstaller](https://pyinstaller.org/) kullanılabilir.

```bash
pip install pyinstaller
pyinstaller --name personeltak --onefile src/personeltak_app.py
```

Komut sonrası `dist/personeltak.exe` dosyası paylaşılmaya hazırdır; aynı yapılandırma dosyası ve standart çalışan listesi konumu ile ağ üzerinde birden fazla kullanıcı tarafından eş zamanlı kullanılabilir.

## Tek Dosya Yapısı

`src/personeltak_app.py` dosyası; konfigürasyon okuma, Excel yükleme, skor hesaplama, rapor oluşturma, kayıt ekleme ve CLI uçlarını tek bir yerde barındırır. `pyproject.toml` içindeki `personeltak` komutu bu dosyanın `main` fonksiyonunu çağırır. Böylece projeyi PyCharm, PyInstaller veya paylaşılan ağ klasörlerinde taşırken ek modül bağımlılığıyla uğraşmanız gerekmez.

## Test Senaryoları

Örnek kabul kriterleri ve test senaryoları proje gereksinimlerinde listelenmiştir. Gerçek veri ile test etmek için örnek Excel dosyasını hazırlayarak CLI komutunu çalıştırmanız yeterlidir.
