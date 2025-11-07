# Tek Komutla PersonelTak Çalıştırma

Aşağıdaki tek satırlık komut, projeyi sıfırdan kurup raporu üretir. Komut, sanal ortam yaratır, bağımlılıkları yükler, örnek yapılandırmayı `config.yaml` olarak kopyalar ve standart çalışan listesini kullanarak raporları `raporlar/` klasörüne yazar.

> **Ön koşullar**: Windows 10/11, PowerShell 5+ veya PowerShell Core, Python 3.10+ kurulu olmalıdır. Çalışan listesi `config.example.yaml` dosyasındaki `employees_path` ile aynı standart ağ dizininde bulunmalıdır.

```powershell
powershell -ExecutionPolicy Bypass -Command "python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt; if (-not (Test-Path 'config.yaml')) {Copy-Item 'config.example.yaml' 'config.yaml'}; python src/personeltak_app.py summarize --config config.yaml --output raporlar"
```

Komutu PowerShell penceresine yapıştırmanız yeterlidir. İlk çalıştırmada `config.yaml` dosyası `config.example.yaml`'dan kopyalanır; daha sonra ihtiyaçlarınıza göre güncelleyebilirsiniz. `employees_path` konumu aynı ağ paylaşımını gösterdiği sürece, uygulama çalışan listesini bu klasörden otomatik okur.

## Linux/Mac Alternatifi

```bash
python -m venv .venv && source .venv/bin/activate \
  && pip install -r requirements.txt \
  && [ -f config.yaml ] || cp config.example.yaml config.yaml \
  && python src/personeltak_app.py summarize --config config.yaml --output raporlar
```

## HTML Panelini Tek Komutla Açmak

Tarayıcı tabanlı paneli başlatmak için rapor komutu yerine `web` alt komutunu kullanabilirsiniz. Örnek PowerShell komutu:

```powershell
powershell -ExecutionPolicy Bypass -Command "python -m venv .venv; .\.venv\Scripts\Activate.ps1; pip install -r requirements.txt; if (-not (Test-Path 'config.yaml')) {Copy-Item 'config.example.yaml' 'config.yaml'}; python src/personeltak_app.py web --config config.yaml --host 0.0.0.0 --port 8050"
```

Linux/macOS için eşdeğer tek satırlık komut:

```bash
python -m venv .venv && source .venv/bin/activate \
  && pip install -r requirements.txt \
  && [ -f config.yaml ] || cp config.example.yaml config.yaml \
  && python src/personeltak_app.py web --config config.yaml --host 0.0.0.0 --port 8050
```

## Tek Komutun Yaptıkları

1. **Sanal ortam oluşturur** (`.venv`).
2. **Bağımlılıkları kurar** (`pandas`, `openpyxl`, `pyyaml`).
3. **Varsayılan yapılandırmayı kopyalar** ve çalışan listesi yolu dahil tüm parametreleri hazırlar.
4. **Raporu üretir**: Excel/CSV/Power BI çıktıları, log dosyası ve kilit yönetimi dahil tek seferde çalışır.

## Uygulamanın `.exe` Olarak Kullanımı

Kurulumdan sonra tek bir komutla `.exe` üretmek için:

```powershell
powershell -ExecutionPolicy Bypass -Command "pip install pyinstaller; pyinstaller --name personeltak --onefile src/personeltak_app.py"
```

Bu komut `dist/personeltak.exe` çıktısını verir; dosyayı paylaşılan klasöre kopyalayıp aynı çalışan listesi ve `config.yaml` ile eşzamanlı kullanabilirsiniz.

## Notlar

- Komutu ihtiyaç duyduğunuz her sunucu veya kullanıcı bilgisayarında çalıştırabilirsiniz.
- PowerShell komutu, tek satırlık olması için `;` ile zincirlenmiştir; dilerseniz adım adım da çalıştırabilirsiniz.
- Raporlar varsayılan olarak `raporlar/` klasörüne yazılır; farklı bir yol kullanmak için komutun sonundaki `--output` parametresini güncelleyin.
