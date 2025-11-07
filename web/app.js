const ROLES = ["Personel", "Şef", "Yönetici"]; // Web arayüzünde desteklenen rol sıralaması

const DEFAULT_CONFIG = { // Python tarafındaki varsayılan ayarların tarayıcı karşılığı
  role_weights: { Personel: 0.2, Şef: 0.4, Yönetici: 0.4 }, // Rol ağırlıkları
  category_weights: { İş: 1.0, Kanaat: 0.7 }, // Kategori ağırlıkları
  tespit_days: 30, // Tespit kriteri için gün aralığı
  missing_threshold: null, // Eksik eşik değeri (isteğe bağlı)
};

const state = { // Uygulama durumunu tutan nesne
  workbook: null, // Yüklenen Excel içeriği
  lastFileName: null, // Son yüklenen dosyanın adı
  asof: todayIso(), // Kullanıcının seçtiği rapor tarihi
};

const workbookInput = document.getElementById("workbook-input"); // Dosya yükleme bileşeni
const asofInput = document.getElementById("asof-input"); // Tarih seçim alanı
const refreshButton = document.getElementById("refresh-button"); // Raporu hesapla düğmesi
const downloadButton = document.getElementById("download-button"); // Excel indir düğmesi
const statusText = document.getElementById("status-text"); // Durum bilgisi metni
const messageArea = document.getElementById("message-area"); // Uyarı/sistem mesajları alanı
const scoresTable = document.getElementById("scores-table"); // Skor tablosu
const missingTable = document.getElementById("missing-table"); // Eksik tablosu
const footerText = document.getElementById("footer-text"); // Alt bilgi metni
const recordForm = document.getElementById("record-form"); // Yeni kayıt formu

asofInput.value = state.asof; // Başlangıçta tarih alanını bugünün tarihiyle doldur

workbookInput.addEventListener("change", handleWorkbookUpload); // Dosya seçildiğinde yükleme işlemini başlat
refreshButton.addEventListener("click", refreshReport); // Raporu hesapla düğmesine tıklandığında hesaplama yap
recordForm.addEventListener("submit", handleRecordSubmit); // Form gönderildiğinde yeni kayıt ekle

function handleWorkbookUpload(event) { // Dosya yükleme işleyicisi
  const file = event.target.files?.[0]; // Kullanıcının seçtiği ilk dosyayı al
  if (!file) { // Dosya seçilmemişse
    return; // İşlemi sonlandır
  }
  statusText.textContent = "Dosya okunuyor..."; // Kullanıcıya süreç bilgisi ver
  const reader = new FileReader(); // Tarayıcı dosyayı okuyabilsin diye FileReader oluştur
  reader.onload = (loadEvent) => { // Dosya başarıyla okunduğunda çalışacak geri çağırım
    try { // Hata yakalamak için korumalı blok aç
      const data = new Uint8Array(loadEvent.target.result); // Okunan veriyi bayt dizisine çevir
      const workbook = XLSX.read(data, { type: "array" }); // SheetJS ile Excel içeriğini oku
      state.workbook = normalizeWorkbook(workbook); // Excel içeriğini uygulama formatına çevir
      state.lastFileName = file.name; // Dosya adını durum bilgisine kaydet
      statusText.textContent = "Dosya başarıyla yüklendi."; // Kullanıcıya başarılı mesajı göster
      footerText.textContent = `Yüklenen dosya: ${file.name}`; // Alt bilgiye dosya adını yaz
      messageArea.innerHTML = renderAlert("Dosya yüklendi. Raporu hesaplamak için \"Raporu Hesapla\" düğmesine basın.", "success"); // Bilgilendirici mesaj ekle
      downloadButton.disabled = false; // Excel indir düğmesini etkinleştir
    } catch (error) { // Yükleme sırasında hata oluşursa
      console.error(error); // Geliştirici konsoluna hatayı yaz
      state.workbook = null; // Mevcut workbook referansını sıfırla
      downloadButton.disabled = true; // İndirme düğmesini pasifleştir
      messageArea.innerHTML = renderAlert(`Dosya okunamadı: ${error.message}`, "error"); // Kullanıcıya hata mesajı göster
      statusText.textContent = "Dosya okunamadı."; // Durum metnini güncelle
      footerText.textContent = "Dosya yüklenmedi."; // Alt bilgiye varsayılan metni yaz
    }
  }; // onload sonu
  reader.onerror = () => { // Dosya okuma hatasında çalışacak geri çağırım
    state.workbook = null; // Workbook referansını temizle
    downloadButton.disabled = true; // İndirme düğmesini kapat
    messageArea.innerHTML = renderAlert("Dosya okunurken bir hata oluştu.", "error"); // Hata mesajı göster
    statusText.textContent = "Dosya okunamadı."; // Durum metnini güncelle
    footerText.textContent = "Dosya yüklenmedi."; // Alt bilgiye varsayılan metni yaz
  }; // onerror sonu
  reader.readAsArrayBuffer(file); // Dosyayı ikili tampon olarak oku
}

function normalizeWorkbook(workbook) { // Excel içeriğini kullanılabilir formata çeviren yardımcı
  const requiredSheets = ["Kriterler", "Calisanlar", "Degerlendirmeler"]; // Gerekli sayfa adları
  const sheetNames = workbook.SheetNames; // Çalışma kitabındaki mevcut sayfalar
  for (const sheet of requiredSheets) { // Her gerekli sayfa için döngü başlat
    if (!sheetNames.includes(sheet)) { // Sayfa eksikse
      throw new Error(`${sheet} sayfası bulunamadı.`); // Kullanıcıya açıklayıcı hata üret
    }
  }
  const criteria = XLSX.utils.sheet_to_json(workbook.Sheets["Kriterler"], { defval: "" }); // Kriterler sayfasını JSON formatına dönüştür
  const employees = XLSX.utils.sheet_to_json(workbook.Sheets["Calisanlar"], { defval: "" }); // Çalışanlar sayfasını JSON formatına dönüştür
  const evaluationsRaw = XLSX.utils.sheet_to_json(workbook.Sheets["Degerlendirmeler"], { defval: "" }); // Değerlendirmeler sayfasını JSON formatına dönüştür
  const evaluations = evaluationsRaw // Değerlendirmeleri normalize etmek için haritalama başlat
    .map((row) => ({ // Her satır için yeni bir nesne oluştur
      Sicil: String(row.Sicil ?? "").trim(), // Sicil alanını temizle
      Rol: String(row.Rol ?? "").trim(), // Rol alanını temizle
      Po: toNumber(row.Po), // Po alanını sayıya çevir
      Puan: toNumber(row.Puan), // Puan alanını sayıya çevir
      Tarih: toDate(row.Tarih), // Tarih alanını Date nesnesine çevir
      HaftaYili: inferIsoWeek(row.HaftaYili, row.Tarih), // Hafta bilgisini hesapla
      Not: String(row.Not ?? "").trim(), // Not alanını temizle
      Period: row.Period ? String(row.Period).trim() : "", // Period alanını koru
    }))
    .filter((row) => row.Sicil && row.Rol && Number.isFinite(row.Po) && Number.isFinite(row.Puan) && row.Tarih); // Eksik alanları olan kayıtları temizle
  return { criteria, employees, evaluations }; // Normalize edilmiş yapıyı döndür
}

function refreshReport() { // Raporu yeniden hesaplayan fonksiyon
  if (!state.workbook) { // Excel yüklenmediyse
    messageArea.innerHTML = renderAlert("Önce PersonelTak çalışma kitabını yükleyin.", "error"); // Kullanıcıyı uyar
    return; // İşlemi sonlandır
  }
  const selectedDate = asofInput.value || todayIso(); // Tarih alanındaki değeri al, boşsa bugünü kullan
  state.asof = selectedDate; // Durum bilgisine tarihi kaydet
  try { // Hesaplama sırasında oluşabilecek hataları yakalamak için blok aç
    const result = computeSummary(state.workbook, DEFAULT_CONFIG, new Date(selectedDate)); // Python mantığını JS ile çalıştır
    renderSummary(result); // Hesaplanan sonuçları arayüze yaz
    statusText.textContent = "Rapor güncellendi."; // Durum metnini başarı mesajıyla güncelle
    messageArea.innerHTML = renderWarnings(result); // Uyarıları veya başarı mesajını göster
  } catch (error) { // Hesaplama başarısız olursa
    console.error(error); // Konsola ayrıntı yaz
    messageArea.innerHTML = renderAlert(`Rapor hesaplanamadı: ${error.message}`, "error"); // Kullanıcıya hata mesajı göster
    statusText.textContent = "Hesaplama hatası."; // Durum metnini güncelle
  }
}

function computeSummary(workbook, config, asofDate) { // Ana raporlama algoritması
  if (!(asofDate instanceof Date) || Number.isNaN(asofDate.valueOf())) { // Tarih geçerli değilse
    throw new Error("Geçerli bir rapor tarihi seçin."); // Kullanıcıya anlamlı hata döndür
  }
  const criteriaMap = new Map(); // Po -> kriter eşleştirmesi için harita oluştur
  workbook.criteria.forEach((criterionRaw) => { // Her kriter satırı için döngü başlat
    const po = toNumber(criterionRaw.Po); // Po değerini sayıya çevir
    if (!Number.isFinite(po)) { // Po geçersizse
      return; // Bu kriteri atla
    }
    criteriaMap.set(po, criterionRaw); // Kriteri haritaya kaydet
  });
  if (criteriaMap.size === 0) { // Hiç kriter bulunamadıysa
    throw new Error("Kriterler sayfası boş görünüyor."); // Kullanıcıyı bilgilendir
  }
  const allowedRoleMap = new Map(); // Po -> izinli roller haritası oluştur
  criteriaMap.forEach((criterion, po) => { // Her kriter için izinli rolleri çıkar
    allowedRoleMap.set(po, allowedRoles(criterion)); // Rol listesi hesapla
  });

  const normalizedEvaluations = workbook.evaluations.filter((row) => criteriaMap.has(row.Po)); // Geçersiz Po içeren kayıtları temizle
  const ignoredEvaluations = workbook.evaluations.length - normalizedEvaluations.length; // Kaç kayıt atıldığını hesapla
  const warnings = []; // Uyarı mesajlarını toplayacak dizi oluştur
  if (ignoredEvaluations > 0) { // Atılan kayıt varsa
    warnings.push(`${ignoredEvaluations} kayıt geçersiz Po nedeniyle yok sayıldı.`); // Uyarı ekle
  }

  const isoWeek = isoWeekString(asofDate); // Seçili tarihin ISO hafta bilgisini hesapla
  const tespitSince = new Date(asofDate.getTime() - config.tespit_days * 24 * 60 * 60 * 1000); // Tespit periyodu başlangıcını belirle

  const scoresRows = []; // Skor satırlarını tutacak dizi
  const missingRows = []; // Eksik kayıt satırlarını tutacak dizi

  workbook.employees.forEach((employee) => { // Her çalışan için döngü başlat
    const sicil = String(employee.Sicil ?? "").trim(); // Çalışan sicilini al
    if (!sicil) { // Sicil yoksa
      return; // Bu satırı atla
    }
    const employeeScores = []; // Çalışanın skor katkılarını tutacak liste
    const employeeWeights = []; // Ağırlıkları tutacak liste
    const employeeMissing = []; // Eksik kayıtları tutacak liste

    criteriaMap.forEach((criterion, po) => { // Her kriter için döngü başlat
      const category = criterion.Kategori ? String(criterion.Kategori).trim() : "İş"; // Kategori bilgisini al
      const period = criterion.Period ? String(criterion.Period).trim() : "Haftalık"; // Period bilgisini al
      const puanMax = toNumber(criterion.PuanMax) || 5; // Maksimum puanı belirle
      const allowedRolesForPo = allowedRoleMap.get(po) ?? []; // Bu Po için izinli rolleri çek
      if (allowedRolesForPo.length === 0) { // Hiç rol yoksa
        return; // Bu kriteri atla
      }
      const roleScores = {}; // Rol bazlı skorları tutan nesne oluştur
      const missingRoles = []; // Eksik roller listesini hazırla

      allowedRolesForPo.forEach((role) => { // Her izinli rol için döngü başlat
        const matches = normalizedEvaluations // Uygun değerlendirmeleri filtrele
          .filter((evaluation) => evaluation.Sicil === sicil && evaluation.Po === po && evaluation.Rol === role); // Sicil, Po ve Rol'e göre eşleştir
        let candidates = matches; // Başlangıçta tüm eşleşmeler adayıdır
        if (period.toLowerCase() === "haftalık") { // Haftalık kriter ise
          candidates = matches.filter((evaluation) => evaluation.HaftaYili === isoWeek); // Aynı ISO haftasında olanları tut
        } else if (period.toLowerCase() === "tespit") { // Tespit kriteri ise
          candidates = matches.filter((evaluation) => evaluation.Tarih >= tespitSince && evaluation.Tarih <= asofDate); // Tarih aralığını uygula
        } else { // Bilinmeyen bir period varsa
          warnings.push(`Po ${po} için bilinmeyen period: ${period}`); // Uyarı ver
          return; // Bu rolü geç
        }
        if (candidates.length === 0) { // Hiç aday yoksa
          return; // Bu rol için skor ekleme
        }
        const lastRecord = candidates.sort((a, b) => a.Tarih - b.Tarih).at(-1); // Tarihi en yeni kaydı bul
        if (!lastRecord) { // Güvenlik: kayıt yoksa
          return; // İşlemi sonlandır
        }
        const clampedScore = clamp(lastRecord.Puan, 0, puanMax); // Puanı alt/üst sınırlar arasında tut
        roleScores[role] = puanMax ? clampedScore / puanMax : 0; // Normalize edilmiş skoru kaydet
      });

      if (Object.keys(roleScores).length === 0) { // Hiç skor elde edilmediyse
        employeeMissing.push({ // Eksik listesine bilgi ekle
          Sicil: sicil, // Sicil bilgisi
          AdSoyad: employee.AdSoyad ?? "", // Ad soyad bilgisi
          Po: po, // Po numarası
          Değerlendirme: criterion.Değerlendirme ?? "", // Kriter açıklaması
          Period: period, // Period bilgisi
          Eksik_Roller: allowedRolesForPo.join(", "), // Eksik roller
        });
        return; // Bu kriter için işleme devam etmeye gerek yok
      }

      let numerator = 0; // Ağırlıklı ortalama pay kısmı
      let denominator = 0; // Ağırlıklı ortalama payda kısmı
      allowedRolesForPo.forEach((role) => { // Her rol için ağırlıkları uygula
        const weight = toNumber(config.role_weights[role]) || 0; // Rol ağırlığını al
        if (role in roleScores) { // Rol için skor varsa
          numerator += weight * roleScores[role]; // Payı artır
          denominator += weight; // Paydayı artır
        } else { // Rol için skor yoksa
          missingRoles.push(role); // Eksik rol listesine ekle
        }
      });

      if (denominator === 0) { // Ağırlıkların toplamı sıfırsa
        employeeMissing.push({ // Eksik listesine kayıt ekle
          Sicil: sicil, // Sicil bilgisi
          AdSoyad: employee.AdSoyad ?? "", // Ad soyad
          Po: po, // Po numarası
          Değerlendirme: criterion.Değerlendirme ?? "", // Kriter açıklaması
          Period: period, // Period bilgisi
          Eksik_Roller: allowedRolesForPo.join(", "), // Eksik roller
        });
        return; // Bu kriterin değerlendirmesini sonlandır
      }

      const categoryWeight = toNumber(config.category_weights[category]) || 1; // Kategori ağırlığını belirle
      employeeScores.push((numerator / denominator) * categoryWeight); // Skor katkısını listeye ekle
      employeeWeights.push(categoryWeight); // Aynı ağırlığı ağırlık listesine ekle

      if (missingRoles.length > 0) { // Eksik roller varsa
        employeeMissing.push({ // Eksik listesini güncelle
          Sicil: sicil, // Sicil bilgisi
          AdSoyad: employee.AdSoyad ?? "", // Ad soyad
          Po: po, // Po numarası
          Değerlendirme: criterion.Değerlendirme ?? "", // Kriter açıklaması
          Period: period, // Period bilgisi
          Eksik_Roller: missingRoles.join(", "), // Eksik roller
        });
      }
    });

    const totalScore = employeeScores.length > 0 && employeeWeights.length > 0 // Skor hesaplanabilir mi kontrol et
      ? (100 * sum(employeeScores)) / sum(employeeWeights) // Toplam skoru hesapla
      : 0; // Veri yoksa sıfır kullan

    scoresRows.push({ // Skor tablosuna satır ekle
      Sicil: sicil, // Sicil numarası
      AdSoyad: employee.AdSoyad ?? "", // Ad soyad
      Departman: employee.Departman ?? "", // Departman
      Unvan: employee.Unvan ?? "", // Unvan
      ToplamSkor: round2(totalScore), // İki basamaklı skor
      Hafta: isoWeek, // ISO hafta bilgisi
    });

    missingRows.push(...employeeMissing); // Eksik kayıtları genel listeye ekle
  });

  let filteredMissing = missingRows; // Eksik kayıtları filtrelemek için başlangıç
  if (Number.isInteger(config.missing_threshold) && config.missing_threshold > 0) { // Eksik eşiği tanımlandıysa
    const counts = countBy(filteredMissing, (row) => row.Sicil); // Sicil bazında kaç eksik olduğunu say
    filteredMissing = filteredMissing.filter((row) => counts.get(row.Sicil) >= config.missing_threshold); // Eşiğin altını el
  }

  return { // Hesaplanan sonuçları döndür
    scores: scoresRows,
    missing: filteredMissing,
    warnings,
    asof: asofDate,
  };
}

function renderSummary(result) { // Hesaplanan sonuçları arayüze yazan fonksiyon
  scoresTable.innerHTML = result.scores.length > 0 ? toTable(result.scores) : "<p class=\"empty\">Skor bulunamadı.</p>"; // Skor tablosunu güncelle
  missingTable.innerHTML = result.missing.length > 0 ? toTable(result.missing) : "<p class=\"empty\">Eksik puanlama yok.</p>"; // Eksik tablosunu güncelle
  footerText.textContent = `Son hesaplama: ${formatDateTime(new Date())}`; // Alt bilgiye zaman damgası yaz
}

function renderWarnings(result) { // Uyarıları kullanıcıya gösteren yardımcı
  if (result.warnings.length === 0) { // Uyarı yoksa
    return renderAlert("Rapor başarıyla güncellendi.", "success"); // Başarı mesajı döndür
  }
  const listItems = result.warnings.map((text) => `<li>${escapeHtml(text)}</li>`).join("\n"); // Uyarıları liste öğesine çevir
  return `<div class=\"alert alert-warning\"><strong>Uyarılar:</strong><ul>${listItems}</ul></div>`; // Uyarı kutusunu oluştur
}

function handleRecordSubmit(event) { // Form gönderme işleyicisi
  event.preventDefault(); // Tarayıcının varsayılan gönderim davranışını engelle
  if (!state.workbook) { // Excel yüklenmemişse
    messageArea.innerHTML = renderAlert("Önce Excel dosyasını yükleyin.", "error"); // Kullanıcıyı uyar
    return; // İşlemi sonlandır
  }
  const formData = new FormData(recordForm); // Form verilerini oku
  const sicil = String(formData.get("sicil") ?? "").trim(); // Sicil değerini al
  const rol = String(formData.get("rol") ?? "").trim(); // Rol değerini al
  const po = toNumber(formData.get("po")); // Po değerini sayıya çevir
  const puan = toNumber(formData.get("puan")); // Puan değerini sayıya çevir
  const tarihInput = String(formData.get("tarih") ?? "").trim(); // Tarih metnini al
  const note = String(formData.get("note") ?? "").trim(); // Not alanını al
  const periodOverride = String(formData.get("period") ?? "").trim(); // Period override değerini al
  if (!sicil || !rol || !Number.isFinite(po) || !Number.isFinite(puan)) { // Gerekli alanlar eksikse
    messageArea.innerHTML = renderAlert("Formu eksiksiz doldurun.", "error"); // Hata mesajı göster
    return; // İşlemi sonlandır
  }
  const tarih = tarihInput ? new Date(tarihInput) : new Date(); // Tarih girdisini Date nesnesine çevir
  if (Number.isNaN(tarih.valueOf())) { // Tarih geçersizse
    messageArea.innerHTML = renderAlert("Tarih alanı geçerli değil.", "error"); // Hata mesajı göster
    return; // İşlemi sonlandır
  }
  const newRecord = { // Yeni değerlendirme kaydını oluştur
    Sicil: sicil,
    Rol: rol,
    Po: po,
    Puan: puan,
    Tarih: tarih,
    HaftaYili: isoWeekString(tarih),
    Not: note,
    Period: periodOverride,
  };
  state.workbook.evaluations.push(newRecord); // Kaydı mevcut değerlendirmelere ekle
  messageArea.innerHTML = renderAlert("Yeni kayıt hafızaya alındı. Excel'e yazdırmak için \"Excel İndir\" butonunu kullanın.", "success"); // Kullanıcıya bilgi ver
  refreshReport(); // Raporu güncelle
  recordForm.reset(); // Formu sıfırla
}

function renderAlert(message, tone) { // Basit uyarı kutusu oluşturan yardımcı
  const toneClass = tone === "error" ? "alert-error" : tone === "warning" ? "alert-warning" : "alert-success"; // Renk sınıfını belirle
  return `<div class=\"alert ${toneClass}\">${escapeHtml(message)}</div>`; // HTML çıktısını döndür
}

function renderAlertText(message, tone) { // Metin tabanlı durum mesajı için yardımcı
  const toneClass = tone === "error" ? "alert-error" : tone === "warning" ? "alert-warning" : "alert-success"; // Sınıf seçimi
  return `<span class=\"${toneClass}\">${escapeHtml(message)}</span>`; // Span öğesi döndür
}

downloadButton.addEventListener("click", () => { // Excel indir düğmesine tıklandığında çalışacak fonksiyon
  if (!state.workbook) { // Excel yoksa
    messageArea.innerHTML = renderAlert("Önce Excel dosyası yükleyin.", "error"); // Kullanıcıyı bilgilendir
    return; // İşlemi sonlandır
  }
  try { // Dışa aktarım sırasında oluşabilecek hataları yakala
    const workbook = XLSX.utils.book_new(); // Yeni çalışma kitabı oluştur
    const criteriaSheet = XLSX.utils.json_to_sheet(state.workbook.criteria); // Kriterler sayfasını oluştur
    const employeesSheet = XLSX.utils.json_to_sheet(state.workbook.employees); // Çalışanlar sayfasını oluştur
    const evaluationsSheet = XLSX.utils.json_to_sheet(state.workbook.evaluations.map((row) => ({ // Değerlendirmeleri Excel'e uygun hale getir
      ...row,
      Tarih: formatDateTime(row.Tarih),
    })));
    XLSX.utils.book_append_sheet(workbook, criteriaSheet, "Kriterler"); // Çalışma kitabına kriterleri ekle
    XLSX.utils.book_append_sheet(workbook, employeesSheet, "Calisanlar"); // Çalışma kitabına çalışanları ekle
    XLSX.utils.book_append_sheet(workbook, evaluationsSheet, "Degerlendirmeler"); // Çalışma kitabına değerlendirmeleri ekle
    const fileName = state.lastFileName ? state.lastFileName.replace(/\.xlsx?$/i, "_web.xlsx") : "personeltak_web.xlsx"; // Dosya adını belirle
    XLSX.writeFile(workbook, fileName); // Dosyayı kullanıcının bilgisayarına indir
    messageArea.innerHTML = renderAlertText(`Excel dosyası \"${fileName}\" olarak indirildi.`, "success"); // Kullanıcıya durum bildir
  } catch (error) { // Hata oluşursa
    console.error(error); // Konsola hatayı yaz
    messageArea.innerHTML = renderAlert(`Excel oluşturulamadı: ${error.message}`, "error"); // Kullanıcıya hata mesajı göster
  }
});

function allowedRoles(criterion) { // Kriter satırından izinli rolleri çıkarır
  return ROLES.filter((role) => String(criterion[role] ?? "").trim().toLowerCase() === "x"); // "x" işareti olan rolleri döndür
}

function toNumber(value) { // Değerleri güvenli şekilde sayıya çevirir
  const parsed = Number(value); // Sayıya dönüştürmeyi dene
  return Number.isFinite(parsed) ? parsed : NaN; // Geçerliyse döndür, değilse NaN ver
}

function toDate(value) { // Excel tarih alanını Date nesnesine çevirir
  if (!value && value !== 0) { // Değer boşsa
    return null; // Tarih yok
  }
  if (value instanceof Date) { // Zaten Date nesnesi ise
    return value; // Olduğu gibi kullan
  }
  if (typeof value === "number") { // Seri numarası şeklindeyse
    return XLSX.SSF.parse_date_code(value) // SheetJS tarih çözümlemesini kullan
      ? new Date(Date.UTC(1899, 11, 30) + value * 86400000) // Excel seri tarihini gerçek tarihe çevir
      : null; // Çözümlenemiyorsa null dön
  }
  const parsed = new Date(value); // Genel amaçlı parse
  return Number.isNaN(parsed.valueOf()) ? null : parsed; // Geçerliyse Date döndür, değilse null
}

function inferIsoWeek(weekValue, dateValue) { // Hafta bilgisi eksikse hesaplayan yardımcı
  const week = String(weekValue ?? "").trim(); // Hafta hücresinin değerini al
  if (week) { // Değer zaten varsa
    return week; // Olduğu gibi döndür
  }
  const date = toDate(dateValue); // Tarihi Date nesnesine çevir
  return date ? isoWeekString(date) : ""; // Tarih geçerliyse ISO hafta hesapla, yoksa boş bırak
}

function isoWeekString(date) { // ISO hafta biçimini döndürür
  const tmp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate())); // Tarihi UTC ortamına taşı
  const dayNum = tmp.getUTCDay() || 7; // Haftanın gününü al (0 ise 7 yap)
  tmp.setUTCDate(tmp.getUTCDate() + 4 - dayNum); // ISO haftasının perşembesine hizala
  const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1)); // Yılın ilk gününü al
  const weekNo = Math.ceil(((tmp - yearStart) / 86400000 + 1) / 7); // Hafta numarasını hesapla
  return `${tmp.getUTCFullYear()}-W${String(weekNo).padStart(2, "0")}`; // ISO hafta metnini döndür
}

function clamp(value, min, max) { // Değeri belirtilen aralıkta kısıtla
  return Math.min(Math.max(value, min), max); // Alt ve üst sınırları uygula
}

function sum(list) { // Liste elemanlarını toplayan yardımcı
  return list.reduce((acc, value) => acc + value, 0); // reduce ile toplam döndür
}

function round2(value) { // Sayıyı iki ondalığa yuvarla
  return Math.round(value * 100) / 100; // Basit yuvarlama uygula
}

function countBy(list, keyFn) { // Listeyi anahtara göre gruplayıp sayan fonksiyon
  const map = new Map(); // Sonuçları tutacak harita
  list.forEach((item) => { // Her öğe için döngü
    const key = keyFn(item); // Anahtarı hesapla
    const current = map.get(key) || 0; // Mevcut sayıyı al
    map.set(key, current + 1); // Sayaç değerini artır
  });
  return map; // Haritayı döndür
}

function toTable(rows) { // Nesne dizisini HTML tablosuna çevirir
  if (rows.length === 0) { // Veri yoksa
    return "<p class=\"empty\">Veri bulunamadı.</p>"; // Boş mesajı döndür
  }
  const headers = Object.keys(rows[0]); // Sütun başlıklarını al
  const thead = `<thead><tr>${headers.map((key) => `<th>${escapeHtml(key)}</th>`).join("")}</tr></thead>`; // Başlık satırını oluştur
  const tbody = `<tbody>${rows
    .map((row) => `<tr>${headers.map((key) => `<td>${escapeHtml(formatCell(row[key]))}</td>`).join("")}</tr>`)
    .join("")}</tbody>`; // Gövde satırlarını üret
  return `<table>${thead}${tbody}</table>`; // Tam tabloyu döndür
}

function formatCell(value) { // Hücre değerini kullanıcıya uygun biçime çevirir
  if (value instanceof Date) { // Değer tarih ise
    return formatDateTime(value); // Tarih formatını kullan
  }
  if (typeof value === "number") { // Sayı ise
    return value.toLocaleString("tr-TR", { maximumFractionDigits: 2 }); // Türkçe biçimlendirme uygula
  }
  return String(value ?? ""); // Diğer durumlarda metne çevir
}

function escapeHtml(value) { // HTML kaçış fonksiyonu
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatDateTime(date) { // Tarihi YYYY-MM-DD HH:MM formatına çevirir
  const year = date.getFullYear(); // Yılı al
  const month = String(date.getMonth() + 1).padStart(2, "0"); // Ayı iki haneli yaz
  const day = String(date.getDate()).padStart(2, "0"); // Günü iki haneli yaz
  const hours = String(date.getHours()).padStart(2, "0"); // Saati iki haneli yaz
  const minutes = String(date.getMinutes()).padStart(2, "0"); // Dakikayı iki haneli yaz
  return `${year}-${month}-${day} ${hours}:${minutes}`; // Formatlanmış tarih döndür
}

function todayIso() { // Bugünün tarihini ISO formatında döndürür
  return new Date().toISOString().slice(0, 10); // YYYY-MM-DD formatını üret
}
