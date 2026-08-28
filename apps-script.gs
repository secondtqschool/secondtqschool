/**
 * نظام حصر ومتابعة غياب الطالبات — المتوسطة التاسعة والستون
 * هذا الكود يُلصق كاملاً داخل محرر Apps Script المرتبط بجدول بيانات Google Sheets.
 * بعد أي تعديل على هذا الكود: Deploy > Manage deployments > تعديل (القلم) > Version: New version > Deploy
 * حتى يعمل التعديل فعليًا على الرابط المنشور (تعديل الكود وحده لا يكفي بدون هذه الخطوة).
 */

const ACCESS_KEY = 'hayat2qschool';
const SHEET_NAME = 'Responses';
const DRIVE_FOLDER_NAME = 'مرفقات غياب الطالبات - الثانية لتحفيظ القرآن';
const ALERT_EMAIL = 'tootaa.b.o.j@gmail.com';

// نفس عتبات لوحة المتابعة بالضبط — بحسب إجمالي أيام الغياب فقط
function statusLevel(total) {
  if (total >= 10) return 'r';
  if (total >= 5)  return 'o';
  if (total >= 3)  return 'y';
  return 'g';
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'reset') {
      if (data.key !== ACCESS_KEY) return jsonResponse({ status: 'unauthorized' });
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      return jsonResponse({ status: 'ok' });
    }

    if (data.action === 'delete') {
      if (data.key !== ACCESS_KEY) return jsonResponse({ status: 'unauthorized' });
      const sheet = getSheet();
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][7] === data.id) {
          sheet.deleteRow(i + 1);
          return jsonResponse({ status: 'ok' });
        }
      }
      return jsonResponse({ status: 'not_found' });
    }

    const sheet = getSheet();

    // منع تسجيل نفس الغياب مرتين لنفس الطالبة في نفس التاريخ
    const existingRows = sheet.getDataRange().getValues().slice(1);
    const isDuplicate = existingRows.some(r =>
      r[4] === data.studentName && String(r[2]) === String(data.grade) &&
      String(r[3]) === String(data.section) && String(r[1]) === String(data.date)
    );
    if (isDuplicate) {
      return jsonResponse({ status: 'duplicate' });
    }

    // حالة الطالبة قبل إضافة هذا السجل (لمعرفة إن كانت هذه أول مرة تصل فيها للأحمر)
    const before = getStudentStats(sheet, data.studentName, data.grade, data.section);
    const beforeLevel = statusLevel(before.total);

    let fileUrl = '';
    if (data.fileBase64 && data.fileName) {
      fileUrl = saveFile(data.fileBase64, data.fileName, data.fileMime);
    }

    const id = Utilities.getUuid();
    sheet.appendRow([
      new Date(),
      data.date || '',
      data.grade || '',
      data.section || '',
      data.studentName || '',
      data.reason || '',
      fileUrl,
      id
    ]);

    // تثبيت التاريخ والفصل كنص صريح، لمنع Google Sheets من تحويلهما تلقائيًا
    // (الفصل قيمته رقم مفرد مثل "١" ويحوّله Sheets تلقائيًا إلى رقم فعلي، ما يكسر أي مقارنة نصية لاحقة)
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat('@STRING@').setValue(data.date || '');
    sheet.getRange(lastRow, 4).setNumberFormat('@STRING@').setValue(data.section || '');

    // إعادة الحساب بعد الإضافة، وإرسال تنبيه إذا كانت هذه أول مرة تصل الطالبة للحالة الحمراء (١٠ أيام غياب فأكثر)
    const after = getStudentStats(sheet, data.studentName, data.grade, data.section);
    const afterLevel = statusLevel(after.total);
    if (afterLevel === 'r' && beforeLevel !== 'r') {
      sendUrgentAlert(data.studentName, data.grade, data.section, after);
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });
  }
}

function doGet(e) {
  const key = e.parameter.key;
  if (key !== ACCESS_KEY) {
    return jsonResponse({ status: 'unauthorized' });
  }
  const sheet = getSheet();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(r => r[4]).map(r => ({
    date: String(r[1]),
    grade: r[2],
    section: r[3],
    studentName: r[4],
    reason: r[5],
    fileUrl: r[6],
    id: r[7]
  }));
  return jsonResponse({ status: 'ok', records: rows });
}

function getStudentStats(sheet, studentName, grade, section) {
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1).filter(r =>
    r[4] === studentName && String(r[2]) === String(grade) && String(r[3]) === String(section)
  );
  const total = rows.length;
  const undocumented = rows.filter(r => !r[6]).length;
  return { total, undocumented };
}

function sendUrgentAlert(studentName, grade, section, stats) {
  const subject = 'تنبيه عاجل: تكرار غياب الطالبة ' + studentName;
  const body =
    'تنبيه من نظام حصر ومتابعة غياب الطالبات — المتوسطة التاسعة والستون\n\n' +
    'وصلت الطالبة (' + studentName + ') — ' + grade + ' فصل ' + section +
    ' إلى مستوى "تحتاج إجراء عاجل" في لوحة المتابعة.\n\n' +
    'إجمالي أيام الغياب: ' + stats.total + '\n' +
    'غياب بدون مرفق داعم: ' + stats.undocumented + '\n\n' +
    'يرجى مراجعة لوحة المتابعة لاتخاذ الإجراء المناسب.';
  MailApp.sendEmail(ALERT_EMAIL, subject, body);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['الوقت', 'التاريخ', 'الصف', 'الفصل', 'اسم الطالبة', 'سبب الغياب', 'رابط المرفق', 'المعرف']);
  }
  return sheet;
}

function saveFile(base64, name, mime) {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mime, name);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
