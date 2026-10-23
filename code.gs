// ==========================================
// PARAMETER APLIKASI (DAPAT DISESUAIKAN)
// ==========================================
const APP_CONFIG = {
  judul: "Ujian Sumatif Akhir Semester (SAS)",
  mapel: "Ilmu Pengetahuan Alam",
  sekolah: "SMPN 01 Konoha",
  kelas: "Kelas 7 / Fase D",
  durasi_menit: 90,
  admin_pass: "admincbt123"
};

// ==========================================
// SETUP & ROUTING
// ==========================================
function doGet() {
  let template = HtmlService.createTemplateFromFile('Index');
  template.config = APP_CONFIG;
  return template.evaluate()
    .setTitle(APP_CONFIG.judul + " - " + APP_CONFIG.sekolah)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getConfig() {
  return APP_CONFIG;
}

// ==========================================
// DATABASE SETUP (JALANKAN SEKALI)
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Master_Soal
  let sheetSoal = ss.getSheetByName("Master_Soal");
  if (!sheetSoal) {
    sheetSoal = ss.insertSheet("Master_Soal");
    sheetSoal.appendRow(["ID_Soal", "Nomor_Soal", "Jenis_Soal", "Teks_Stimulus", "Gambar_Stimulus_URL", "Teks_Soal", "Gambar_Soal_URL", "Pilihan_A", "Pilihan_B", "Pilihan_C", "Pilihan_D", "Kunci_Jawaban", "Pembahasan", "Bobot"]);
    sheetSoal.setFrozenRows(1);
    
    // Sampel Soal
    sheetSoal.appendRow(["S01", 1, "PG", "Bumi adalah planet ketiga dari Matahari...", "", "Apa nama planet ketiga dari Matahari?", "", "Venus", "Mars", "Bumi", "Jupiter", "C", "Bumi adalah planet ketiga di tata surya.", 10]);
    sheetSoal.appendRow(["S02", 2, "PG", "", "", "Hewan yang bernapas dengan insang adalah...", "https://placehold.co/400x200/e2e8f0/475569?text=Gambar+Hewan", "Ayam", "Kucing", "Ikan", "Kambing", "C", "Ikan hidup di air dan bernapas menggunakan insang.", 10]);
    sheetSoal.appendRow(["S03", 3, "PG", "Perhatikan tabel fotosintesis berikut...", "", "Zat apa yang dihasilkan dari proses fotosintesis?", "", "Oksigen dan Glukosa", "Karbon Dioksida dan Air", "Oksigen dan Air", "Karbon Dioksida dan Glukosa", "A", "Fotosintesis menghasilkan O2 dan C6H12O6.", 20]);
  }

  // 2. Setup Master_Siswa
  let sheetSiswa = ss.getSheetByName("Master_Siswa");
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet("Master_Siswa");
    sheetSiswa.appendRow(["Nama_Siswa", "Kelas", "Token_Akses"]);
    sheetSiswa.setFrozenRows(1);
    
    // Sampel Siswa
    const siswa = [
      ["Uzumaki Naruto", "7A", "TOKEN1"],
      ["Uchiha Sasuke", "7A", "TOKEN2"],
      ["Haruno Sakura", "7B", "TOKEN3"],
      ["Hatake Kakashi", "7B", "TOKEN4"],
      ["Nara Shikamaru", "7C", "TOKEN5"]
    ];
    siswa.forEach(s => sheetSiswa.appendRow(s));
  }

  // 3. Setup Hasil_Ujian
  let sheetHasil = ss.getSheetByName("Hasil_Ujian");
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet("Hasil_Ujian");
    sheetHasil.appendRow(["Timestamp", "Nama_Siswa", "Kelas", "Nilai_Akhir", "Total_Benar", "Total_Salah", "Pelanggaran_Tab", "Detail_Jawaban_JSON"]);
    sheetHasil.setFrozenRows(1);
  }
  
  return "Database berhasil disetup!";
}

// ==========================================
// FUNGSI AUTHENTICATION
// ==========================================
function verifyStudentLogin(nama, token) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Siswa");
  if (!sheet) throw new Error("Database belum disetup.");
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === nama.toLowerCase() && data[i][2].toString() === token) {
      return { success: true, nama: data[i][0], kelas: data[i][1] };
    }
  }
  return { success: false, message: "Nama atau Token salah!" };
}

function verifyAdminLogin(password) {
  if (password === APP_CONFIG.admin_pass) {
    return { success: true };
  }
  return { success: false, message: "Password Admin salah!" };
}

// ==========================================
// FUNGSI SISWA (AMBIL SOAL TANPA KUNCI)
// ==========================================
function getQuestionsForStudent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Soal");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let questions = [];
  
  for (let i = 1; i < data.length; i++) {
    let q = {};
    for (let j = 0; j < headers.length; j++) {
      // KEAMANAN: Jangan kirim Kunci & Pembahasan ke browser siswa
      if (headers[j] !== "Kunci_Jawaban" && headers[j] !== "Pembahasan") {
        q[headers[j]] = data[i][j];
      }
    }
    questions.push(q);
  }
  // Sort by Nomor_Soal
  questions.sort((a, b) => a.Nomor_Soal - b.Nomor_Soal);
  return questions;
}

// ==========================================
// FUNGSI PENILAIAN & SUBMIT (AUTO-SCORING)
// ==========================================
function submitStudentExam(studentData, answers, violations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSoal = ss.getSheetByName("Master_Soal");
  const sheetHasil = ss.getSheetByName("Hasil_Ujian");
  
  // Ambil Kunci Jawaban
  const dataSoal = sheetSoal.getDataRange().getValues();
  const headers = dataSoal[0];
  const idxId = headers.indexOf("ID_Soal");
  const idxKunci = headers.indexOf("Kunci_Jawaban");
  const idxBobot = headers.indexOf("Bobot");
  
  let mapKunci = {};
  let mapBobot = {};
  let totalBobotMax = 0;
  
  for (let i = 1; i < dataSoal.length; i++) {
    let id = dataSoal[i][idxId];
    mapKunci[id] = dataSoal[i][idxKunci].toString().trim().toUpperCase();
    let bobot = parseFloat(dataSoal[i][idxBobot]) || 10;
    mapBobot[id] = bobot;
    totalBobotMax += bobot;
  }
  
  // Hitung Skor
  let totalBenar = 0;
  let totalSalah = 0;
  let skorDiperoleh = 0;
  
  for (let id in answers) {
    let jwb = answers[id].toString().trim().toUpperCase();
    if (jwb === mapKunci[id]) {
      totalBenar++;
      skorDiperoleh += mapBobot[id];
    } else {
      totalSalah++;
    }
  }
  
  // Skala Nilai Akhir (0-100)
  let nilaiAkhir = totalBobotMax > 0 ? (skorDiperoleh / totalBobotMax) * 100 : 0;
  nilaiAkhir = Math.round(nilaiAkhir * 100) / 100;
  
  // Simpan ke Sheet Hasil
  const timestamp = new Date();
  sheetHasil.appendRow([
    timestamp, 
    studentData.nama, 
    studentData.kelas, 
    nilaiAkhir, 
    totalBenar, 
    totalSalah, 
    violations, 
    JSON.stringify(answers)
  ]);
  
  return {
    success: true,
    nilaiAkhir: nilaiAkhir,
    benar: totalBenar,
    salah: totalSalah
  };
}

// ==========================================
// FUNGSI ADMIN
// ==========================================
function saveSingleQuestion(qData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master_Soal");
  const newId = "S" + new Date().getTime().toString().substr(-6);
  sheet.appendRow([
    newId,
    qData.Nomor_Soal,
    qData.Jenis_Soal,
    qData.Teks_Stimulus,
    qData.Gambar_Stimulus_URL,
    qData.Teks_Soal,
    qData.Gambar_Soal_URL,
    qData.Pilihan_A,
    qData.Pilihan_B,
    qData.Pilihan_C,
    qData.Pilihan_D,
    qData.Kunci_Jawaban,
    qData.Pembahasan,
    qData.Bobot
  ]);
  return { success: true, message: "Soal berhasil disimpan!" };
}

function getExamResultsForAdmin() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hasil_Ujian");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let results = [];
  for(let i = 1; i < data.length; i++) {
    results.push({
      waktu: data[i][0],
      nama: data[i][1],
      kelas: data[i][2],
      nilai: data[i][3],
      benar: data[i][4],
      salah: data[i][5],
      pelanggaran: data[i][6]
    });
  }
  return results.reverse(); // Terbaru di atas
}
