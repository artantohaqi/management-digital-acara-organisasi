/**
 * Konfigurasi ID Form dan ID Google Calendar
 */
const FORM_ID = "xxxx"; // Ganti dengan ID Form
const CALENDAR_ID = "xxxx@group.calendar.google.com"; // Ganti dengan ID Kalender
const ADMIN_EMAIL = "xxxx@gmail.com"; // Ganti dengan email admin

/**
 * Menangani submit dari Google Form
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(lastRow, 1, 1, 14).getValues()[0];

  const email = data[1];
  const editAcara = data[3];
  const jenisAcara = data[4];
  const namaAcara = data[5];
  const tanggalMulai = new Date(data[6]);
  const jamMulaiRaw = data[7];
  const deskripsiAcara = data[8];
  const tanggalSelesai = data[9] ? new Date(data[9]) : new Date(tanggalMulai);
  const jamSelesaiRaw = data[10];

  const jamMulai = parseTime(jamMulaiRaw);
  Logger.log(`Jam Mulai sebelum setHours: ${jamMulai.hours}:${jamMulai.minutes}`);
  tanggalMulai.setHours(jamMulai.hours, jamMulai.minutes, 0); // Pastikan menit diterapkan
  Logger.log(`Jam Mulai setelah setHours: ${tanggalMulai.getHours()}:${tanggalMulai.getMinutes()}`);

  let jamSelesai;
  if (!jamSelesaiRaw) {
    // Jika jam selesai kosong, tambah 2 jam dari jam mulai
    jamSelesai = {
      hours: (jamMulai.hours + 2) % 24,
      minutes: jamMulai.minutes
    };
    Logger.log(`Jam Selesai dihitung: ${jamSelesai.hours}:${jamSelesai.minutes} (dari Jam Mulai + 2 jam)`);
  } else {
    jamSelesai = parseTime(jamSelesaiRaw);
    Logger.log(`Jam Selesai sebelum setHours: ${jamSelesai.hours}:${jamSelesai.minutes}`);
  }
  tanggalSelesai.setHours(jamSelesai.hours, jamSelesai.minutes, 0); // Pastikan menit diterapkan
  Logger.log(`Jam Selesai setelah setHours: ${tanggalSelesai.getHours()}:${tanggalSelesai.getMinutes()}`);

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const eventTitle = `${jenisAcara} - ${namaAcara}`;
  const eventId = Utilities.getUuid();
  let iCalUID;

  // Jika mengedit acara lama
  if (editAcara) {
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][12] === editAcara) {
        iCalUID = allData[i][13];

        // Kosongkan Dropdown Label & iCalUID lama
        sheet.getRange(i + 1, 13).clearContent(); // Kolom M
        sheet.getRange(i + 1, 14).clearContent(); // Kolom N

        try {
          const oldEvent = calendar.getEventById(iCalUID);
          if (oldEvent) oldEvent.deleteEvent();
        } catch (e) {
          Logger.log("Gagal menghapus acara lama: " + e.toString());
        }
        break;
      }
    }
  }

  // Simpan ID Acara Baru
  sheet.getRange(lastRow, 12).setValue(eventId); // Kolom L

  try {
    const event = calendar.createEvent(eventTitle, tanggalMulai, tanggalSelesai, {
      description: deskripsiAcara,
      guests: email,
      sendInvites: true,
    });

    event.removeAllReminders();
    event.addPopupReminder(1440); // 1 hari
    event.addPopupReminder(300);  // 5 jam
    event.addPopupReminder(180);  // 3 jam

    sheet.getRange(lastRow, 14).setValue(event.getId()); // Kolom N

    MailApp.sendEmail(email, `Acara berhasil dibuat`, `Acara '${eventTitle}' berhasil dijadwalkan.`);
    MailApp.sendEmail(ADMIN_EMAIL, `Acara baru oleh ${email}`, eventTitle);
  } catch (e) {
    MailApp.sendEmail(email, `Gagal menjadwalkan acara`, e.toString());
  }

  updateDropdownLabelColumn();
  updateFormDropdownOptions();
}

/**
 * Parsing waktu seperti "07:00 AM"
 */
function parseTime(timeStr) {
  // Konversi ke string dan log input awal
  const inputStr = String(timeStr || '').trim();
  Logger.log(`Input timeStr: ${inputStr} (Tipe: ${typeof timeStr})`);

  let hours = 0;
  let minutes = 0;
  let period = "AM";

  // Coba ekstrak waktu dari string panjang (misalnya, dari Date)
  const timeMatch = inputStr.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s?(AM|PM)/i); // Menangani format seperti 11:11:00 AM
  if (timeMatch) {
    hours = parseInt(timeMatch[1], 10) || 0;
    minutes = parseInt(timeMatch[2], 10) || 0; // Ambil menit dari timeMatch[2]
    period = (timeMatch[3] || "AM").toUpperCase();
    Logger.log(`Ekstraksi berhasil: ${hours}:${minutes} ${period}`);
  } else {
    // Pisah berdasarkan pemisah standar
    const parts = inputStr.split(/[: ]/).filter(part => part && part.length);
    Logger.log(`Parts setelah split: ${parts}`);

    if (parts.length >= 2) {
      hours = parseInt(parts[0], 10) || 0;
      minutes = parseInt(parts[1], 10) || 0;
      if (parts.length >= 3 && ["AM", "PM"].includes(parts[2].toUpperCase())) {
        period = parts[2].toUpperCase();
      } else {
        Logger.log(`Periode tidak valid atau hilang: ${parts}`);
      }
    } else {
      Logger.log(`Pencocokan gagal: Format waktu tidak valid. Input: ${inputStr}. Harus dalam format HH:MM AM/PM`);
      return { hours: 0, minutes: 0 };
    }
  }

  // Konversi 12-jam ke 24-jam
  if (period === "PM" && hours !== 12) hours += 12;
  if (period === "AM" && hours === 12) hours = 0;

  // Validasi rentang
  hours = Math.min(Math.max(hours, 0), 23);
  minutes = Math.min(Math.max(minutes, 0), 59);

  Logger.log(`Hasil parsing: ${hours}:${minutes} ${period} (24h: ${hours})`);
  return { hours, minutes };
}

/**
 * Update kolom Dropdown Label berdasarkan data valid
 */
function updateDropdownLabelColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const jenis = data[i][4];
    const nama = data[i][5];
    const tanggal = data[i][6];
    const jam = data[i][7]; // Ambil langsung dari kolom H (Jam Mulai)
    const uid = data[i][13];

    if (nama && uid) {
      const tanggalFormatted = Utilities.formatDate(new Date(tanggal), "GMT+07:00", "dd/MM/yyyy");
      const label = `${jenis} - ${nama} | ${tanggalFormatted} | ${jam}`; // Gunakan nilai asli dari kolom H
      sheet.getRange(i + 1, 13).setValue(label); // Kolom M
    } else {
      sheet.getRange(i + 1, 13).setValue(""); // Kosongkan jika tidak valid
    }
  }
}

/**
 * Update pilihan dropdown di Google Form (untuk Edit Acara)
 */
function updateFormDropdownOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const labels = sheet.getRange(2, 13, sheet.getLastRow() - 1).getValues()
    .map(row => row[0])
    .filter(label => label && label.trim() !== "");

  const form = FormApp.openById(FORM_ID);
  const items = form.getItems(FormApp.ItemType.LIST);

  for (const item of items) {
    if (item.getTitle().toLowerCase().includes("edit acara")) {
      item.asListItem().setChoiceValues(labels);
    }
  }
}

/**
 * Cek acara yang akan dimulai & kirim pengingat
 */
function checkAndSendReminders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  for (let i = 1; i < data.length; i++) {
    const title = `${data[i][4]} - ${data[i][5]}`;
    const tanggal = new Date(data[i][6]);
    const jam = data[i][7];
    const id = data[i][13];
    if (!id || !jam) continue;

    const parsed = parseTime(jam);
    tanggal.setHours(parsed.hours, parsed.minutes, 0);

    const diffMins = Math.floor((tanggal - now) / 60000);
    if ([1440, 300, 180].includes(diffMins)) {
      try {
        const event = calendar.getEventById(id);
        const reminderText = `Pengingat: '${title}' akan dimulai pada:\n` +
          Utilities.formatDate(tanggal, "GMT+07:00", "EEEE, dd MMMM yyyy hh:mm a");
        event.getGuestList().forEach(g =>
          MailApp.sendEmail(g.getEmail(), `Reminder: ${title}`, reminderText));
      } catch (e) {
        Logger.log("Reminder error: " + e.toString());
      }
    }
  }

  updateFormDropdownOptions();
}

/**
 * Jalankan otomatis saat Spreadsheet dibuka
 */
function runOnOpen() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === "onFormSubmit" || t.getHandlerFunction() === "checkAndSendReminders") {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger("checkAndSendReminders")
    .timeBased()
    .everyMinutes(1)
    .create();

  updateDropdownLabelColumn();
  updateFormDropdownOptions();
  Logger.log("Triggers dan dropdown diperbarui saat onOpen.");
}
