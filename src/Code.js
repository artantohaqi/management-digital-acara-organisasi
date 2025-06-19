/**
 * Konfigurasi ID Form dan ID Google Calendar
 */
const FORM_ID = "xxxxx"; // Ganti dengan ID Form
const CALENDAR_ID = "xxxxx@group.calendar.google.com"; // Ganti dengan ID Kalender
const ADMIN_EMAIL = "xxxxx@gmail.com"; // Ganti dengan email admin

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
  tanggalMulai.setHours(jamMulai.hours, jamMulai.minutes, 0);

  const jamSelesai = parseTime(jamSelesaiRaw || "12:00 PM");
  tanggalSelesai.setHours(jamSelesai.hours, jamSelesai.minutes, 0);

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
  const parts = timeStr.toString().split(/[: ]/);
  let hours = parseInt(parts[0]) || 0;
  const minutes = parseInt(parts[1]) || 0;
  const period = (parts[2] || "AM").toUpperCase();

  if (period === "PM" && hours !== 12) hours += 12;
  if (period === "AM" && hours === 12) hours = 0;

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
    const jam = data[i][7];
    const uid = data[i][13];

    if (nama && uid) {
      const tanggalFormatted = Utilities.formatDate(new Date(tanggal), "GMT+07:00", "dd/MM/yyyy");
      const label = `${jenis} - ${nama} | ${tanggalFormatted} | ${jam}`;
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
