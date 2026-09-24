/**
 * Perbaikan: toggle "Approval Berjenjang" kembali mati setelah disimpan.
 *
 * Penyebab: sheet Pengaturan dibuat SEBELUM kolom ApprovalSpkAktif &
 * JumlahTingkatApproval ditambahkan ke SCHEMA. getSheet_() tidak pernah
 * memperbarui header sheet yang sudah ada, jadi:
 *   - updateBaris_ MENULIS nilai ke kolom K/L (posisi menurut SCHEMA), tapi
 *   - bacaSemua_ MEMBACA berdasarkan label header baris 1 -- K1/L1 kosong,
 *     sehingga baris.ApprovalSpkAktif selalu undefined -> terbaca "mati".
 * Pengaturan tidak ada di SHEET_PERLU_MIGRASI_STRUKTUR_ (kolom barunya
 * selalu ditambah di akhir), jadi aman diperbaiki otomatis cukup dengan
 * melengkapi label header baris 1.
 */

// === 1) TAMBAHKAN fungsi baru ini di BAGIAN 3: Db.gs (mis. tepat di bawah
//        fungsi perbaikiHeaderSemua) ===

/**
 * Melengkapi label header baris 1 sebuah sheet yang kolom barunya hanya
 * DITAMBAH DI AKHIR (bukan disisipkan di tengah) -- aman, hanya menulis
 * label yang masih kosong, tidak pernah memindahkan data. Kalau pola
 * headernya tidak cocok (ada label yang berbeda), sheet TIDAK disentuh.
 * @return {boolean} true kalau header diperbaiki.
 */
function selaraskanHeaderKolomAkhir_(namaSheet) {
  if (SHEET_PERLU_MIGRASI_STRUKTUR_[namaSheet]) return false;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet) return false;
  var header = SCHEMA[namaSheet];
  var lebar = Math.max(sheet.getLastColumn(), header.length);
  var headerSaatIni = sheet.getRange(1, 1, 1, lebar).getValues()[0];
  var perlu = false;
  for (var i = 0; i < lebar; i++) {
    var lama = headerSaatIni[i];
    var baru = i < header.length ? header[i] : '';
    if (lama === baru) continue;
    if (lama === '' && baru !== '') { perlu = true; continue; }
    return false;
  }
  if (perlu) sheet.getRange(1, 1, 1, header.length).setValues([header]);
  return perlu;
}

// === 2) GANTI fungsi pastikanPengaturanSiap_ (BAGIAN 5) dengan ini ===

function pastikanPengaturanSiap_() {
  return withLock_(function () {
    getSheet_(CONFIG.SHEET_PENGATURAN);
    selaraskanHeaderKolomAkhir_(CONFIG.SHEET_PENGATURAN);
    var ada = cariSatu_(CONFIG.SHEET_PENGATURAN, 'PengaturanId', CONFIG.ID_PENGATURAN);
    if (ada) return false;
    var sekarang = sekarangIso_();
    tambahBaris_(CONFIG.SHEET_PENGATURAN, {
      PengaturanId: CONFIG.ID_PENGATURAN,
      NamaPT: '',
      Alamat: '',
      LogoUrl: '',
      DiubahOleh: '',
      DiubahPada: sekarang
    });
    return true;
  });
}

// === 3) GANTI pengaturanNotifikasiAmbil_ & pengaturanApprovalAmbil_
//        (BAGIAN 11B) dengan ini ===

function pengaturanNotifikasiAmbil_() {
  selaraskanHeaderKolomAkhir_(CONFIG.SHEET_PENGATURAN);
  var baris = cariSatu_(CONFIG.SHEET_PENGATURAN, 'PengaturanId', CONFIG.ID_PENGATURAN);
  return { emailNotifikasi: (baris && baris.EmailNotifikasi) || '' };
}

function pengaturanApprovalAmbil_() {
  selaraskanHeaderKolomAkhir_(CONFIG.SHEET_PENGATURAN);
  var baris = cariSatu_(CONFIG.SHEET_PENGATURAN, 'PengaturanId', CONFIG.ID_PENGATURAN);
  var diatur = !!(baris && (baris.ApprovalSpkAktif === true || baris.ApprovalSpkAktif === 'TRUE'));
  return {
    // aktif  = benar-benar berlaku (sudah memperhitungkan paket).
    // diatur = nilai yang tersimpan di Pengaturan (untuk mengisi toggle form).
    aktif: diatur && fiturDiizinkanAman_('approval_berjenjang'),
    diatur: diatur,
    jumlahTingkat: (baris && Number(baris.JumlahTingkatApproval)) || 1
  };
}

// === 4) Di akhir pengaturanSimpan_ (BAGIAN 11B), GANTI baris
//          return pengaturanAmbil_();
//        dengan: ===
//
//    var hasil = pengaturanAmbil_();
//    hasil.approval = pengaturanApprovalAmbil_();
//    return hasil;

// ==========================================================================
// === 5) JS.html -- di bukaPengaturanPerusahaan() ===
// ==========================================================================
// Toggle diisi dari nilai TERSIMPAN (pa.diatur), bukan pa.aktif -- pa.aktif
// ikut mati kalau paket terbaca bukan Premium, sehingga toggle tampil mati
// walau pengaturannya sebenarnya tersimpan menyala.
//
// a) GANTI baris fallback:
//      apiCall('pengaturanApprovalAmbil', {}).catch(function () { return { aktif: false, jumlahTingkat: 1 }; })
//    dengan:
//      apiCall('pengaturanApprovalAmbil', {}).catch(function () { return { aktif: false, diatur: false, jumlahTingkat: 1 }; })
//
// b) Tepat di bawah baris:
//      var pa = hasil[2] || { aktif: false, jumlahTingkat: 1 };
//    TAMBAHKAN:
//      var approvalDiatur = pa.diatur !== undefined ? pa.diatur : pa.aktif;
//
// c) GANTI potongan html switch:
//      '<input class="form-check-input" type="checkbox" role="switch" id="swApprovalAktif"' + (pa.aktif ? ' checked' : '') + '>' +
//      '<label class="form-check-label small" for="swApprovalAktif">Aktifkan Approval Berjenjang</label>' +
//      '</div>' +
//    dengan:
//      '<input class="form-check-input" type="checkbox" role="switch" id="swApprovalAktif"' + (approvalDiatur ? ' checked' : '') + '>' +
//      '<label class="form-check-label small" for="swApprovalAktif">Aktifkan Approval Berjenjang</label>' +
//      '</div>' +
//      (approvalDiatur && !pa.aktif
//        ? '<div class="small text-warning mb-2">Tersimpan aktif, tapi belum berlaku karena paket lisensi saat ini bukan Premium.</div>'
//        : '') +
