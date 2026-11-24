
const SPREADSHEET_ID = '1abNTbDhWpXuIFQ2dtRqm2hKRhol6TWhVXcReJGhDtrU';
const DATE_SHEET_ID = '1N4BThkx1HXKcM6DqyIHg64q_bIUodkJcb8_KUF6tnAw';
const FOLDER_ID = '1b_qnc-fmDyHWE_Mzy4GYssaL0sKI09Xy';
//const TEACHERS_SPREADSHEET_ID = '1jyGmsfOFkPr05SLjm8M-HXhGXQzcHIWAYlP7bVv4TVE';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('المجتمع المهني - ربعي بن عامر')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getData() {
  try {
    var result = getSheetData();
    return result;
  } catch (error) {
    return getEmptyData('⚠️ خطأ');
  }
}

function getEmptyData(message) {
  return {
    attendees: [],
    locations: [],
    topics: [],
    topicStatus: {},
    specialties: [],
    leaders: [],
    dateInfo: { success: false, hijri: message, day: message },
    topicNumbers: {}
  };
}

// ====== جلب جميع البيانات (مع أكواد الحاضرين) ======
function getSheetData() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    var attendeesSheet = sheet.getSheetByName('الحاضرون');
    var leadersSheet = sheet.getSheetByName('الرؤساء');
    
    var attendees = [];
    var attendeeCodes = {};
    var locations = [];
    var topics = [];
    var topicStatus = {};
    var specialties = [];
    var leaders = [];
    var leaderCodes = {};
    var topicNumbers = {};
    
    // جلب الحاضرون وأكوادهم
    if (attendeesSheet) {
      var attendeeData = attendeesSheet.getRange('B2:C').getValues();
      attendeeData.forEach(function(row) {
        var name = row[0] ? row[0].toString().trim() : '';
        var code = row[1] ? row[1].toString().trim() : '';
        if (name) {
          attendees.push(name);
          if (code) {
            attendeeCodes[name] = code;
          }
        }
      });
    }
    
    // جلب الرؤساء
    if (leadersSheet) {
      var leaderData = leadersSheet.getRange('B2:C').getValues();
      leaderData.forEach(function(row) {
        var name = row[0] ? row[0].toString().trim() : '';
        var code = row[1] ? row[1].toString().trim() : '';
        if (name) {
          leaders.push(name);
          if (code) {
            leaderCodes[name] = code;
          }
        }
      });
    }
    
    // جلب باقي البيانات
    if (dataSheet) {
      var locationData = dataSheet.getRange('F2:F').getValues();
      locations = locationData.flat().filter(function(l) { return l && l.toString().trim() !== ''; }).map(function(l) { return l.toString().trim(); });
      
      var specialtyData = dataSheet.getRange('K2:K').getValues();
      specialties = specialtyData.flat().filter(function(s) { return s && s.toString().trim() !== ''; }).map(function(s) { return s.toString().trim(); });
      
      var topicData = dataSheet.getRange('H2:H').getValues();
      var topicNames = topicData.flat().filter(function(t) { return t && t.toString().trim() !== ''; }).map(function(t) { return t.toString().trim(); });
      
      var numberData = dataSheet.getRange('G2:G').getValues();
      var numbers = numberData.flat().filter(function(n) { return n && n.toString().trim() !== ''; }).map(function(n) { return n.toString().trim(); });
      
      var dateData = dataSheet.getRange('I2:J').getValues();
      var now = new Date();
      now.setHours(0, 0, 0, 0);
      
      topicNames.forEach(function(topic, index) {
        var num = (index < numbers.length && numbers[index]) ? numbers[index] : '-';
        var topicWithNumber = '[' + num + '] ' + topic;
        topics.push(topicWithNumber);
        topicNumbers[topicWithNumber] = num;
        
        var statusText = '⚠️ بدون تاريخ';
        var isActive = false;
        var start = '-';
        var end = '-';
        
        if (dateData && index < dateData.length && dateData[index] && dateData[index][0] && dateData[index][1]) {
          try {
            var startDate = new Date(dateData[index][0]);
            var endDate = new Date(dateData[index][1]);
            start = formatDateDisplay(startDate);
            end = formatDateDisplay(endDate);
            
            if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
              if (now >= startDate && now <= endDate) {
                statusText = '✅ مفعل';
                isActive = true;
              } else if (now < startDate) {
                var daysLeft = Math.ceil((startDate.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
                statusText = '⏳ باقي ' + daysLeft + ' يوم' + (daysLeft > 1 ? 'اً' : '') + ' للبدء';
              } else if (now > endDate) {
                statusText = '❌ منتهي';
              }
            }
          } catch(e) {}
        }
        
        topicStatus[topicWithNumber] = {
          start: start,
          end: end,
          active: isActive,
          statusText: statusText,
          number: num
        };
      });
    }
    
    var dateInfo = getTodayDate();
    
    return {
      attendees: attendees || [],
      attendeeCodes: attendeeCodes || {},
      locations: locations || [],
      topics: topics || [],
      topicStatus: topicStatus || {},
      specialties: specialties || [],
      leaders: leaders || [],
      leaderCodes: leaderCodes || {},
      dateInfo: dateInfo,
      topicNumbers: topicNumbers || {}
    };
    
  } catch (error) {
    return getEmptyData('⚠️ خطأ في جلب البيانات');
  }
}

function getTodayDate() {
  try {
    var sheet = SpreadsheetApp.openById(DATE_SHEET_ID);
    var dateSheet = sheet.getSheetByName('التاريخ');
    if (!dateSheet) return { success: false, hijri: '⚠️ التاريخ غير موجود', day: '⚠️ التاريخ غير موجود' };
    
    var data = dateSheet.getRange('A:C').getValues();
    var today = new Date();
    var todayYear = today.getFullYear();
    var todayMonth = today.getMonth() + 1;
    var todayDay = today.getDate();
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        try {
          var cellDate = null;
          if (row[0] instanceof Date) {
            cellDate = row[0];
          } else if (typeof row[0] === 'string' || typeof row[0] === 'number') {
            cellDate = parseDateFromSheet(row[0].toString().trim());
          }
          if (cellDate && !isNaN(cellDate.getTime())) {
            if (cellDate.getFullYear() === todayYear && cellDate.getMonth() === today.getMonth() && cellDate.getDate() === todayDay) {
              return { success: true, hijri: row[1]?.toString().trim() || '⚠️ التاريخ غير موجود', day: row[2]?.toString().trim() || '⚠️ التاريخ غير موجود', gregorian: todayYear + '/' + todayMonth + '/' + todayDay };
            }
          }
        } catch(e) {}
      }
    }
    return { success: false, hijri: '⚠️ التاريخ غير موجود', day: '⚠️ التاريخ غير موجود', gregorian: todayYear + '/' + todayMonth + '/' + todayDay };
  } catch (error) {
    return { success: false, hijri: '⚠️ خطأ في جلب التاريخ', day: '⚠️ خطأ في جلب التاريخ' };
  }
}

function parseDateFromSheet(dateStr) {
  try {
    if (!dateStr) return null;
    if (dateStr instanceof Date) return dateStr;
    var str = dateStr.toString().trim();
    var dateObj = new Date(str);
    if (!isNaN(dateObj.getTime())) return dateObj;
    var cleaned = str.replace(/[-–—]/g, '/').replace(/[^0-9\/]/g, '');
    var parts = cleaned.split('/');
    if (parts.length !== 3) return null;
    var year, month, day;
    if (parts[0].length === 4) { year = parseInt(parts[0]); month = parseInt(parts[1]); day = parseInt(parts[2]); }
    else if (parts[2].length === 4) { year = parseInt(parts[2]); month = parseInt(parts[1]); day = parseInt(parts[0]); }
    else { year = parseInt(parts[0]); month = parseInt(parts[1]); day = parseInt(parts[2]); }
    if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
    return new Date(year, month - 1, day);
  } catch(e) { return null; }
}

function formatDateDisplay(date) {
  if (!date || isNaN(date.getTime())) return '-';
  var day = date.getDate().toString().padStart(2, '0');
  var month = (date.getMonth() + 1).toString().padStart(2, '0');
  var year = date.getFullYear();
  return day + '/' + month + '/' + year;
}

function verifyCode(code) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var leadersSheet = sheet.getSheetByName('الرؤساء');
    if (!leadersSheet) return { success: false, message: 'شيت الرؤساء غير موجود' };
    if (code === '1234') return { success: true, name: 'مدير النظام', isAdmin: true, message: 'تم التحقق كمدير' };
    var data = leadersSheet.getRange('B:C').getValues();
    for (var i = 0; i < data.length; i++) {
      var name = data[i][0]?.toString().trim() || '';
      var codeValue = data[i][1]?.toString().trim() || '';
      if (codeValue === code) return { success: true, name: name, isAdmin: false, message: 'تم التحقق بنجاح' };
    }
    return { success: false, message: '❌ الكود غير صحيح' };
  } catch(error) { return { success: false, message: error.message }; }
}

// ====== إضافة موضوع جديد (مع إضافة مؤقت واتساب) ======
function addNewTopic(topicData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    
    if (!dataSheet) throw new Error('شيت البيانات غير موجود');
    
    // ====== العثور على أول صف فارغ في العمود H ======
    var hColumn = dataSheet.getRange('H:H').getValues();
    var newRow = -1;
    for (var i = 1; i < hColumn.length; i++) {
      if (!hColumn[i][0] || hColumn[i][0].toString().trim() === '') {
        newRow = i + 1;
        break;
      }
    }
    if (newRow === -1) newRow = hColumn.length + 1;
    
    // ====== العثور على آخر رقم في العمود G ======
    var gColumn = dataSheet.getRange('G:G').getValues();
    var lastNumber = 0;
    for (var j = 0; j < gColumn.length; j++) {
      if (gColumn[j][0] && gColumn[j][0].toString().trim() !== '') {
        var num = parseInt(gColumn[j][0].toString().trim());
        if (!isNaN(num) && num > lastNumber) lastNumber = num;
      }
    }
    
    var newNumber = lastNumber + 1;
    
    // ====== كتابة البيانات في الأعمدة G:J ======
    dataSheet.getRange('G' + newRow).setValue(newNumber);
    dataSheet.getRange('H' + newRow).setValue(topicData.name);
    dataSheet.getRange('I' + newRow).setValue(topicData.startDate);
    dataSheet.getRange('J' + newRow).setValue(topicData.endDate);
    
    // ====== جدولة رسالة واتساب ======
    try {
      var scheduleResult = scheduleWhatsAppForTopic({
        name: topicData.name,
        startDate: topicData.startDate,
        specialty: topicData.specialty || 'غير محدد',
        departmentHead: topicData.departmentHead || 'غير محدد'
      });
      Logger.log('📤 نتيجة جدولة واتساب: ' + JSON.stringify(scheduleResult));
    } catch (e) {
      Logger.log('⚠️ خطأ في جدولة واتساب: ' + e.toString());
    }
    
    // ====== تحديث شيت المتابعة ======
    try {
      syncTrackerData();
    } catch (e) {
      Logger.log('⚠️ خطأ في تحديث شيت المتابعة: ' + e.toString());
    }
    
    return {
      success: true,
      message: 'تم إضافة الموضوع بنجاح',
      topicWithNumber: '[' + newNumber + '] ' + topicData.name,
      number: newNumber
    };
    
  } catch (error) {
    throw new Error('فشل في إضافة الموضوع: ' + error.message);
  }
}
// ====== الحصول على رابط مجلد التقارير ======
function getReportsFolderUrl() {
  try {
    var folder = DriveApp.getFolderById(FOLDER_ID);
    return {
      success: true,
      url: folder.getUrl(),
      name: folder.getName(),
      id: FOLDER_ID
    };
  } catch (error) {
    return {
      success: false,
      message: 'خطأ في جلب رابط المجلد: ' + error.message
    };
  }
}
// ====== جدولة رسالة واتساب ======
function scheduleWhatsAppForTopic(topicData) {
  try {
    // التحقق من وجود الدوال المطلوبة
    if (typeof getDepartmentHeadsNumbers !== 'function') {
      return { success: false, message: 'دوال واتساب غير متوفرة' };
    }
    
    var startDate = new Date(topicData.startDate);
    var now = new Date();
    
    if (startDate.getTime() < now.getTime()) {
      return { success: false, message: 'تاريخ البدء في الماضي' };
    }
    
    var leaders = getDepartmentHeadsNumbers();
    
    if (leaders.length === 0) {
      return { success: false, message: 'لا يوجد رؤساء أقسام' };
    }
    
    // تخزين بيانات المؤقت
    var triggerData = {
      topicName: topicData.name,
      startDate: topicData.startDate,
      specialty: topicData.specialty || 'غير محدد',
      departmentHead: topicData.departmentHead || 'غير محدد',
      leaders: leaders
    };
    
    // حذف المؤقتات السابقة
    deleteTopicTrigger(topicData.name);
    
    // إنشاء مؤقت جديد
    var trigger = ScriptApp.newTrigger('sendWhatsAppFromTrigger')
      .timeBased()
      .at(startDate)
      .create();
    
    triggerData.triggerId = trigger.getUniqueId();
    
    // حفظ بيانات المؤقت
    saveTriggerData(triggerData);
    
    return {
      success: true,
      message: 'تم جدولة رسالة واتساب',
      triggerId: trigger.getUniqueId(),
      leadersCount: leaders.length
    };
    
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ====== حفظ بيانات المؤقت ======
function saveTriggerData(triggerData) {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = 'trigger_' + triggerData.triggerId;
    props.setProperty(key, JSON.stringify(triggerData));
  } catch(e) {}
}

// ====== حذف مؤقتات الموضوع ======
function deleteTopicTrigger(topicName) {
  try {
    var props = PropertiesService.getScriptProperties();
    var keys = props.getKeys();
    
    keys.forEach(function(key) {
      if (key.startsWith('trigger_')) {
        try {
          var data = props.getProperty(key);
          if (data) {
            var parsed = JSON.parse(data);
            if (parsed.topicName === topicName) {
              props.deleteProperty(key);
            }
          }
        } catch(e) {}
      }
    });
  } catch(e) {}
}

// ====== الحصول على أرقام رؤساء الأقسام ======
function getDepartmentHeadsNumbers() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var attendeesSheet = sheet.getSheetByName('الحاضرون');
    var leadersSheet = sheet.getSheetByName('الرؤساء');
    
    if (!attendeesSheet || !leadersSheet) return [];
    
    // جلب أسماء الرؤساء
    var leaderData = leadersSheet.getRange('B2:B').getValues();
    var leaderNames = leaderData.flat()
      .filter(function(name) { return name && name.toString().trim() !== ''; })
      .map(function(name) { return name.toString().trim(); });
    
    // جلب أرقام الجوالات من شيت الحاضرون
    var attendeeData = attendeesSheet.getRange('B2:E').getValues();
    var leaders = [];
    
    attendeeData.forEach(function(row) {
      var name = row[0] ? row[0].toString().trim() : '';
      var phone = row[3] ? row[3].toString().trim() : '';
      if (name && phone && leaderNames.includes(name)) {
        leaders.push({
          name: name,
          phone: phone
        });
      }
    });
    
    return leaders;
    
  } catch(e) {
    return [];
  }
}

// ====== إنشاء شيت المتابعة مباشرة (بدون syncTrackerData) ======
function createTrackerSheetDirect() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    
    if (!trackerSheet) {
      trackerSheet = sheet.insertSheet('متابعة_التنفيذ');
    }
    
    var allTopics = getAllTopicsWithNumbers();
    var leaders = getLeadersList();
    
    var headers = ['#', 'رئيس القسم'];
    allTopics.forEach(function(t) {
      headers.push('[' + t.number + ']');
    });
    headers.push('المجموع');
    
    trackerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    trackerSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    trackerSheet.getRange(1, 1, 1, headers.length).setBackground('#f0f4ff');
    trackerSheet.getRange(1, 1, 1, headers.length).setFontColor('#000000');
    trackerSheet.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
    
    trackerSheet.setColumnWidth(1, 50);
    trackerSheet.setColumnWidth(2, 200);
    for (var i = 3; i <= headers.length; i++) {
      trackerSheet.setColumnWidth(i, 60);
    }
    
    var lastRow = trackerSheet.getLastRow();
    if (lastRow > 1) {
      trackerSheet.deleteRows(2, lastRow - 1);
    }
    
    if (leaders.length > 0) {
      var rows = [];
      var serial = 1;
      leaders.forEach(function(leader) {
        var row = [serial, leader];
        for (var i = 0; i < allTopics.length; i++) {
          row.push('☐');
        }
        row.push('0/' + allTopics.length);
        rows.push(row);
        serial++;
      });
      trackerSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    trackerSheet.setFrozenRows(1);
    
  } catch (error) {
    // تجاهل الأخطاء
  }
}

// ====== حذف موضوع (طريقة آمنة - نسخ البيانات يدوياً) ======
function deleteTopic(topicName) {
  try {
    Logger.log('🗑️ حذف الموضوع: ' + topicName);
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    
    if (!dataSheet) throw new Error('شيت البيانات غير موجود');
    
    // تنظيف اسم الموضوع
    var cleanName = topicName.replace(/^\[\d+\]\s*/, '');
    var hColumn = dataSheet.getRange('H:H').getValues();
    var rowToDelete = -1;
    
    for (var i = 0; i < hColumn.length; i++) {
      if (hColumn[i][0] && hColumn[i][0].toString().trim() === cleanName) {
        rowToDelete = i + 1;
        break;
      }
    }
    
    if (rowToDelete === -1) throw new Error('الموضوع غير موجود');
    
    // ====== جلب جميع البيانات من الأعمدة G:J ======
    var lastRow = dataSheet.getLastRow();
    var dataRange = dataSheet.getRange('G' + rowToDelete + ':J' + lastRow);
    var data = dataRange.getValues();
    
    // ====== إزالة الصف الأول (المطلوب حذفه) ======
    var newData = [];
    for (var i = 1; i < data.length; i++) {
      newData.push(data[i]);
    }
    
    // ====== مسح الأعمدة G:J بالكامل ======
    var clearRange = dataSheet.getRange('G' + rowToDelete + ':J' + lastRow);
    clearRange.clearContent();
    
    // ====== كتابة البيانات الجديدة (نقل للأعلى) ======
    if (newData.length > 0) {
      var writeRange = dataSheet.getRange('G' + rowToDelete + ':J' + (lastRow - 1));
      writeRange.setValues(newData);
    }
    
    Logger.log('✅ تم حذف الموضوع ونقل ما تحته للأعلى');
    
    // ====== إعادة الترقيم التلقائي ======
    var renumberResult = renumberTopics();
    Logger.log('📊 نتيجة إعادة الترقيم: ' + JSON.stringify(renumberResult));
    
    // ====== تحديث شيت المتابعة ======
    try {
      syncTrackerData();
    } catch (e) {
      Logger.log('⚠️ خطأ في تحديث شيت المتابعة: ' + e.toString());
    }
    
    return {
      success: true,
      message: 'تم حذف الموضوع وإعادة الترقيم',
      renumberResult: renumberResult
    };
    
  } catch (error) {
    throw new Error('فشل في حذف الموضوع: ' + error.message);
  }
}

// ====== Generate AI Content (نسخة مع تصحيح الأخطاء) ======
// ====== Generate AI Content (مع تحديد عدد الأحرف) ======
function generateAIContent(topic, specialty) {
  try {
    var topicClean = topic.replace(/^\[\d+\]\s*/, '');
    
    if (!specialty || specialty === '' || specialty === 'اختر التخصص...') {
      specialty = 'التربية';
    }
    
    var apiKey = '';
    
    // ====== بناء Prompt مع تحديد عدد الأحرف ======
    var prompt = `أنت خبير تربوي متخصص في تدريس "${specialty}". 
    
بناءً على موضوع الاجتماع: "${topicClean}"
قم بتوليد محتوى تعليمي مخصص لمعلمي "${specialty}" فقط.

**تعليمات مهمة - عدد الأحرف المطلوب:**
- الفئة المستهدفة: الحد الأقصى 130 حرف (مع المسافات)
- الهدف: الحد الأقصى 130 حرف (مع المسافات)
- خطة العمل (كل خطوة): الحد الأقصى 185 حرف (مع المسافات)
- المخرجات (كل مخرج): الحد الأقصى 130 حرف (مع المسافات)
- مدة التنفيذ (كل مدة): الحد الأقصى 60 حرف (مع المسافات)

**المطلوب JSON فقط:**
{
  "targetAudience": "نص لا يتجاوز 130 حرف",
  "objective": "نص لا يتجاوز 130 حرف",
  "workPlan": [
    "نص لا يتجاوز 185 حرف",
    "نص لا يتجاوز 185 حرف",
    "نص لا يتجاوز 185 حرف",
    "نص لا يتجاوز 185 حرف"
  ],
  "outputs": [
    "نص لا يتجاوز 130 حرف",
    "نص لا يتجاوز 130 حرف",
    "نص لا يتجاوز 130 حرف",
    "نص لا يتجاوز 130 حرف"
  ],
  "durations": [
    "نص لا يتجاوز 60 حرف",
    "نص لا يتجاوز 60 حرف",
    "نص لا يتجاوز 60 حرف",
    "نص لا يتجاوز 60 حرف"
  ]
}`;

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
    var payload = { 
      contents: [{ 
        parts: [{ text: prompt }] 
      }] 
    };
    
    var options = { 
      method: 'POST', 
      headers: { 'Content-Type': 'application/json' }, 
      payload: JSON.stringify(payload), 
      muteHttpExceptions: true 
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      return getFallbackContentWithLimit(topicClean, specialty);
    }
    
    var result = JSON.parse(response.getContentText());
    var generatedText = '';
    
    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      generatedText = result.candidates[0].content.parts[0].text;
    } else {
      return getFallbackContentWithLimit(topicClean, specialty);
    }
    
    var jsonMatch = generatedText.match(/\{[\s\S]*\}/);
    var content = jsonMatch ? JSON.parse(jsonMatch[0]) : null;
    
    if (content) {
      // ====== تطبيق الحدود على المحتوى ======
      var finalContent = {
        targetAudience: limitText(content.targetAudience || getDefaultTargetAudience(specialty, topicClean), 130),
        objective: limitText(content.objective || getDefaultObjective(specialty, topicClean), 130),
        workPlan: [
          limitText(content.workPlan?.[0] || getDefaultWorkPlan(specialty, topicClean)[0], 185),
          limitText(content.workPlan?.[1] || getDefaultWorkPlan(specialty, topicClean)[1], 185),
          limitText(content.workPlan?.[2] || getDefaultWorkPlan(specialty, topicClean)[2], 185),
          limitText(content.workPlan?.[3] || getDefaultWorkPlan(specialty, topicClean)[3], 185)
        ],
        outputs: [
          limitText(content.outputs?.[0] || getDefaultOutputs(specialty, topicClean)[0], 130),
          limitText(content.outputs?.[1] || getDefaultOutputs(specialty, topicClean)[1], 130),
          limitText(content.outputs?.[2] || getDefaultOutputs(specialty, topicClean)[2], 130),
          limitText(content.outputs?.[3] || getDefaultOutputs(specialty, topicClean)[3], 130)
        ],
        durations: [
          limitText(content.durations?.[0] || getDefaultDurations(specialty, topicClean)[0], 60),
          limitText(content.durations?.[1] || getDefaultDurations(specialty, topicClean)[1], 60),
          limitText(content.durations?.[2] || getDefaultDurations(specialty, topicClean)[2], 60),
          limitText(content.durations?.[3] || getDefaultDurations(specialty, topicClean)[3], 60)
        ]
      };
      
      return {
        success: true,
        content: finalContent
      };
    }
    
    return getFallbackContentWithLimit(topicClean, specialty);
    
  } catch (error) {
    var topicClean = topic.replace(/^\[\d+\]\s*/, '');
    if (!specialty || specialty === '' || specialty === 'اختر التخصص...') {
      specialty = 'التربية';
    }
    return getFallbackContentWithLimit(topicClean, specialty);
  }
}

// ====== دالة تقطيع النص إلى عدد محدد من الأحرف ======
function limitText(text, maxLength) {
  if (!text) return '';
  if (text.length <= maxLength) return text;
  
  // ====== قص النص عند آخر مسافة قبل الحد الأقصى ======
  var truncated = text.substring(0, maxLength);
  var lastSpace = truncated.lastIndexOf(' ');
  
  if (lastSpace > 0) {
    return truncated.substring(0, lastSpace) + '...';
  }
  return truncated + '...';
}

// ====== المحتوى الاحتياطي مع الحدود ======
function getFallbackContentWithLimit(topic, specialty) {
  return {
    success: true,
    content: {
      targetAudience: limitText(getDefaultTargetAudience(specialty, topic), 130),
      objective: limitText(getDefaultObjective(specialty, topic), 130),
      workPlan: [
        limitText(getDefaultWorkPlan(specialty, topic)[0], 185),
        limitText(getDefaultWorkPlan(specialty, topic)[1], 185),
        limitText(getDefaultWorkPlan(specialty, topic)[2], 185),
        limitText(getDefaultWorkPlan(specialty, topic)[3], 185)
      ],
      outputs: [
        limitText(getDefaultOutputs(specialty, topic)[0], 130),
        limitText(getDefaultOutputs(specialty, topic)[1], 130),
        limitText(getDefaultOutputs(specialty, topic)[2], 130),
        limitText(getDefaultOutputs(specialty, topic)[3], 130)
      ],
      durations: [
        limitText(getDefaultDurations(specialty, topic)[0], 60),
        limitText(getDefaultDurations(specialty, topic)[1], 60),
        limitText(getDefaultDurations(specialty, topic)[2], 60),
        limitText(getDefaultDurations(specialty, topic)[3], 60)
      ]
    },
    _fallback: true
  };
}




// ====== اختبار الاتصال بـ Gemini API ======
function testGeminiAPI() {
  try {
    var apiKey = '';
    var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;
    
    var options = {
      method: 'GET',
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      var data = JSON.parse(response.getContentText());
      Logger.log('✅ الاتصال بـ Gemini API ناجح');
      Logger.log('📊 النماذج المتاحة: ' + JSON.stringify(data));
      return {
        success: true,
        message: 'الاتصال بـ Gemini API ناجح',
        models: data.models || []
      };
    } else {
      Logger.log('❌ فشل الاتصال بـ Gemini API: ' + responseCode);
      return {
        success: false,
        message: 'فشل الاتصال: ' + responseCode
      };
    }
  } catch (error) {
    Logger.log('❌ خطأ في اختبار API: ' + error.toString());
    return {
      success: false,
      message: error.message
    };
  }
}

// ====== دوال مساعدة للقيم الافتراضية حسب التخصص ======

function getDefaultTargetAudience(specialty, topic) {
  var audiences = {
    'اللغة العربية': 'معلمو اللغة العربية في المراحل التعليمية المختلفة المهتمون بتطوير مهاراتهم في "' + topic + '"',
    'الرياضيات': 'معلمو الرياضيات في المراحل التعليمية المختلفة المهتمون بتطوير استراتيجياتهم في "' + topic + '"',
    'العلوم': 'معلمو العلوم في المراحل التعليمية المختلفة المهتمون بتطوير مهاراتهم في "' + topic + '"',
    'الدراسات الإسلامية': 'معلمو الدراسات الإسلامية في المراحل التعليمية المختلفة المهتمون بتطوير أساليبهم في "' + topic + '"',
    'اللغة الإنجليزية': 'معلمو اللغة الإنجليزية في المراحل التعليمية المختلفة المهتمون بتطوير مهاراتهم في "' + topic + '"',
    'التاريخ': 'معلمو التاريخ في المراحل التعليمية المختلفة المهتمون بتطوير استراتيجياتهم في "' + topic + '"',
    'الجغرافيا': 'معلمو الجغرافيا في المراحل التعليمية المختلفة المهتمون بتطوير مهاراتهم في "' + topic + '"',
    'التربية': 'المعلمون في مختلف التخصصات المهتمون بتطوير مهاراتهم في "' + topic + '"'
  };
  return audiences[specialty] || 'معلمي ' + specialty + ' المهتمين بـ "' + topic + '"';
}

function getDefaultObjective(specialty, topic) {
  var objectives = {
    'اللغة العربية': 'تعزيز كفاءات معلمي اللغة العربية في إنتاج محتوى رقمي جذاب يُثري العملية التعليمية ويُحسّن مهارات الطلاب اللغوية',
    'الرياضيات': 'تمكين معلمي الرياضيات من استخدام المحتوى الرقمي الجذاب لتبسيط المفاهيم الرياضية وتعزيز الفهم العميق لدى الطلاب',
    'العلوم': 'تطوير قدرات معلمي العلوم في تصميم محتوى رقمي تفاعلي يُعزز الاستكشاف العلمي والتفكير النقدي لدى الطلاب',
    'الدراسات الإسلامية': 'إثراء مهارات معلمي الدراسات الإسلامية في إنتاج محتوى رقمي يعزز القيم الإسلامية ويربط الطلاب بتعاليم دينهم',
    'اللغة الإنجليزية': 'تطوير كفاءات معلمي اللغة الإنجليزية في استخدام المحتوى الرقمي لتحسين مهارات التواصل باللغة الإنجليزية لدى الطلاب',
    'التاريخ': 'تمكين معلمي التاريخ من استخدام المحتوى الرقمي لجعل الأحداث التاريخية أكثر واقعية وتفاعلية للطلاب',
    'الجغرافيا': 'تعزيز قدرات معلمي الجغرافيا في تصميم محتوى رقمي يعرض الظواهر الجغرافية بشكل تفاعلي وجذاب',
    'التربية': 'تعزيز كفاءات المعلمين في إنتاج محتوى رقمي جذاب يُحسّن العملية التعليمية'
  };
  return objectives[specialty] || 'تعزيز كفاءات معلمي ' + specialty + ' في "' + topic + '"';
}

function getDefaultWorkPlan(specialty, topic) {
  var plans = {
    'اللغة العربية': [
      'تصميم أنشطة رقمية لتنمية مهارات القراءة والكتابة في اللغة العربية',
      'إنتاج محتوى رقمي تفاعلي (فيديوهات، إنفوجرافيك، عروض) لتدريس قواعد اللغة العربية',
      'تطبيق استراتيجيات التعلم النشط في تدريس اللغة العربية باستخدام الأدوات الرقمية',
      'تقييم أثر المحتوى الرقمي على تحصيل الطلاب في اللغة العربية وتحسينه'
    ],
    'الرياضيات': [
      'تصميم محاكيات رقمية لتبسيط المفاهيم الرياضية المجردة',
      'إنتاج فيديوهات تعليمية تفاعلية لحل المسائل الرياضية خطوة بخطوة',
      'تطبيق استراتيجيات حل المشكلات الرياضية باستخدام التطبيقات الرقمية',
      'قياس مدى تحسن أداء الطلاب في الرياضيات باستخدام المحتوى الرقمي'
    ],
    'العلوم': [
      'تصميم تجارب علمية افتراضية باستخدام المحاكاة الرقمية',
      'إنتاج محتوى رقمي يشرح الظواهر العلمية بطريقة تفاعلية وجذابة',
      'تطبيق استراتيجيات الاستقصاء العلمي باستخدام الأدوات الرقمية',
      'تقييم فعالية المحتوى الرقمي في تنمية المهارات العلمية للطلاب'
    ],
    'الدراسات الإسلامية': [
      'تصميم محتوى رقمي تفاعلي لتعليم أحكام التلاوة والتفسير',
      'إنتاج أنشطة رقمية لتعزيز القيم الإسلامية والأخلاقية لدى الطلاب',
      'تطبيق استراتيجيات التدريس الرقمي في تدريس الفقه والسيرة النبوية',
      'تقييم أثر المحتوى الرقمي على تحصيل الطلاب في الدراسات الإسلامية'
    ],
    'اللغة الإنجليزية': [
      'تصميم محتوى رقمي لتنمية مهارات الاستماع والمحادثة باللغة الإنجليزية',
      'إنتاج فيديوهات تفاعلية لتعليم القواعد والمفردات الإنجليزية',
      'تطبيق استراتيجيات التواصل الشفهي والكتابي باستخدام الأدوات الرقمية',
      'تقييم تحسن مهارات اللغة الإنجليزية لدى الطلاب باستخدام المحتوى الرقمي'
    ],
    'التاريخ': [
      'تصميم محاكيات رقمية للأحداث التاريخية لتجربة واقعية للطلاب',
      'إنتاج محتوى رقمي (خرائط تفاعلية، جداول زمنية) لتوضيح التسلسل التاريخي',
      'تطبيق استراتيجيات التفكير التاريخي باستخدام المصادر الرقمية',
      'تقييم مدى استيعاب الطلاب للأحداث التاريخية عبر المحتوى الرقمي'
    ],
    'الجغرافيا': [
      'تصميم خرائط تفاعلية رقمية لاستكشاف الظواهر الجغرافية',
      'إنتاج محتوى رقمي (صور، فيديوهات، نماذج ثلاثية الأبعاد) للظواهر الطبيعية',
      'تطبيق استراتيجيات الاستكشاف الجغرافي باستخدام نظم المعلومات الجغرافية',
      'تقييم فهم الطلاب للظواهر الجغرافية عبر المحتوى الرقمي'
    ]
  };
  return plans[specialty] || getDefaultPlans(specialty, topic);
}

function getDefaultPlans(specialty, topic) {
  return [
    'تطوير برنامج تدريبي حول "' + topic + '" لمعلمي ' + specialty,
    'تنفيذ ورش عمل تفاعلية لمعلمي ' + specialty,
    'تطبيق استراتيجيات "' + topic + '" في تدريس ' + specialty,
    'تقييم النتائج وتحسينها لمعلمي ' + specialty
  ];
}

function getDefaultOutputs(specialty, topic) {
  var outputs = {
    'اللغة العربية': [
      'دليل إرشادي لإنتاج محتوى رقمي في اللغة العربية',
      'مكتبة من الأنشطة الرقمية لمهارات اللغة العربية',
      'تطبيقات عملية لاستراتيجيات التدريس الرقمي في اللغة العربية',
      'خطة تطوير مهني لمعلمي اللغة العربية في المجال الرقمي'
    ],
    'الرياضيات': [
      'دليل إرشادي لإنتاج محتوى رقمي في الرياضيات',
      'مكتبة من المحاكيات الرقمية للمفاهيم الرياضية',
      'تطبيقات عملية لاستراتيجيات حل المشكلات الرقمية',
      'خطة تطوير مهني لمعلمي الرياضيات في المجال الرقمي'
    ],
    'العلوم': [
      'دليل إرشادي لإنتاج محتوى رقمي في العلوم',
      'مكتبة من التجارب الافتراضية والاستقصاءات الرقمية',
      'تطبيقات عملية لاستراتيجيات الاستقصاء العلمي الرقمي',
      'خطة تطوير مهني لمعلمي العلوم في المجال الرقمي'
    ],
    'الدراسات الإسلامية': [
      'دليل إرشادي لإنتاج محتوى رقمي في الدراسات الإسلامية',
      'مكتبة من الأنشطة الرقمية لتعزيز القيم الإسلامية',
      'تطبيقات عملية لاستراتيجيات التدريس الرقمي في الدراسات الإسلامية',
      'خطة تطوير مهني لمعلمي الدراسات الإسلامية في المجال الرقمي'
    ],
    'اللغة الإنجليزية': [
      'دليل إرشادي لإنتاج محتوى رقمي في اللغة الإنجليزية',
      'مكتبة من الأنشطة الرقمية لمهارات اللغة الإنجليزية',
      'تطبيقات عملية لاستراتيجيات التواصل الرقمي باللغة الإنجليزية',
      'خطة تطوير مهني لمعلمي اللغة الإنجليزية في المجال الرقمي'
    ]
  };
  return outputs[specialty] || getDefaultOutputsGeneric(specialty, topic);
}

function getDefaultOutputsGeneric(specialty, topic) {
  return [
    'دليل إرشادي حول "' + topic + '" لمعلمي ' + specialty,
    'تطبيقات عملية لاستراتيجيات "' + topic + '" في ' + specialty,
    'تقرير تحليلي عن "' + topic + '" في ' + specialty,
    'خطة تطوير مهني في "' + topic + '" لمعلمي ' + specialty
  ];
}

function getDefaultDurations(specialty, topic) {
  return [
    'الأسبوع الأول: التخطيط والإعداد لمعلمي ' + specialty,
    'الأسبوع الثاني: التنفيذ والتطبيق في ' + specialty,
    'الأسبوع الثالث: المتابعة والتقييم لمعلمي ' + specialty,
    'الأسبوع الرابع: التعديل والتحسين في ' + specialty
  ];
}

function getSaveFolder() {
  try { return DriveApp.getFolderById(FOLDER_ID); } catch(e) { return DriveApp.createFolder('تقارير المجتمع المهني'); }
}

// ====== حفظ وتصدير PDF (نسخة معدلة) ======
function saveAndExportPDF(data) {
  try {
    var leaderName = data.departmentHead || 'غير محدد';
    var meetingTopic = data.meetingTopic || 'غير محدد';
    var meetingNumber = data.meetingNumber || 'غير محدد';
    var specialty = data.specialty || 'غير محدد';
    
    if (meetingTopic && typeof meetingTopic === 'string') {
      if (meetingTopic.match(/^\[\d+\]\s*/)) {
        meetingTopic = meetingTopic.replace(/^\[\d+\]\s*/, '');
      }
      if (meetingTopic.match(/\[\d+\]/)) {
        meetingTopic = meetingTopic.replace(/\[\d+\]\s*/, '');
      }
    }
    
    if (meetingNumber === 'غير محدد' && data.meetingTopicFull) {
      var match = data.meetingTopicFull.match(/\[(\d+)\]/);
      if (match) {
        meetingNumber = match[1];
      }
    }
    
    data.meetingTopic = meetingTopic;
    data.meetingNumber = meetingNumber;
    
    var PRESENTATION_ID = '16fTFZmn6NdoRyZgZut8OzscQ5YBq3EeDdirvcn3av98';
    var originalFile = DriveApp.getFileById(PRESENTATION_ID);
    var tempFileName = 'تقرير مؤقت - ' + new Date().toISOString();
    var tempFile = originalFile.makeCopy(tempFileName);
    var tempFileId = tempFile.getId();
    var tempPresentation = SlidesApp.openById(tempFileId);
    var slides = tempPresentation.getSlides();
    
    if (slides.length > 1) {
      for (var i = slides.length - 1; i > 0; i--) {
        slides[i].remove();
      }
    }
    
    var slide = tempPresentation.getSlides()[0];
    replaceVariablesInSlide(slide, data);
    tempPresentation.saveAndClose();
    
    var pdfBlob = tempFile.getAs('application/pdf');
    var dateStr = new Date().toISOString().split('T')[0];
    var fileName = meetingTopic + ' - ' + specialty + ' - ' + dateStr + '.pdf';
    fileName = fileName.replace(/[\\/:"*?<>|]/g, '');
    pdfBlob.setName(fileName);
    
    var folder = getSaveFolder();
    var pdfFile = folder.createFile(pdfBlob);
    tempFile.setTrashed(true);
    
    // تسجيل التنفيذ
    try {
      updateLeaderExecution(leaderName, meetingTopic);
    } catch(e) {
      // تجاهل أخطاء التسجيل
    }
    
    return {
      success: true,
      message: 'تم تصدير PDF بنجاح',
      pdfUrl: pdfFile.getUrl(),
      pdfId: pdfFile.getId(),
      fileName: fileName,
      folderId: FOLDER_ID,
      meetingTopic: meetingTopic,
      meetingNumber: meetingNumber,
      specialty: specialty,
      executedBy: leaderName,
      executionRecorded: true
    };
    
  } catch (error) {
    throw new Error('فشل في تصدير PDF: ' + error.message);
  }
}

function replaceVariablesInSlide(slide, data) {
  try {
    var meetingNumber = data.meetingNumber || 'غير محدد';
    var meetingTopic = data.meetingTopic || 'غير محدد';
    if (meetingNumber === 'غير محدد' && data.meetingTopicFull) {
      var match = data.meetingTopicFull.match(/\[(\d+)\]/);
      if (match) meetingNumber = match[1];
    }
    if (meetingTopic && typeof meetingTopic === 'string') {
      if (meetingTopic.match(/^\[\d+\]\s*/)) meetingTopic = meetingTopic.replace(/^\[\d+\]\s*/, '');
      if (meetingTopic.match(/\[\d+\]/)) meetingTopic = meetingTopic.replace(/\[\d+\]\s*/, '');
    }
    
    var attendeesList = [];
    if (data.attendeesList && Array.isArray(data.attendeesList)) attendeesList = data.attendeesList;
    else if (data.attendees && typeof data.attendees === 'string') {
      attendeesList = data.attendees.split(',').map(function(a) { return a.trim(); }).filter(function(a) { return a; });
    }
    
    var first8Attendees = attendeesList.slice(0, 8);
    var attendees8Text = '';
    first8Attendees.forEach(function(name, index) { attendees8Text += (index + 1) + '- ' + name + '\n'; });
    attendees8Text = attendees8Text.trim();
    if (first8Attendees.length === 0) { attendees8Text = 'لا يوجد حاضرون'; }
    else if (first8Attendees.length < 8) { for (var i = first8Attendees.length; i < 8; i++) { attendees8Text += '\n' + (i + 1) + '- '; } }
    var attendeesCount = attendeesList.length;
    
    var now = new Date();
    var timeZone = 'Asia/Riyadh';
    var currentTime = Utilities.formatDate(now, timeZone, 'hh:mm');
    var futureTime = new Date(now.getTime() + 45 * 60000);
    var futureTimeFormatted = Utilities.formatDate(futureTime, timeZone, 'hh:mm');
    var dateInfo = getTodayDate();
    var hijriDate = dateInfo.success ? dateInfo.hijri : '⚠️ التاريخ غير موجود';
    var dayName = dateInfo.success ? dateInfo.day : '⚠️ اليوم غير موجود';
    
    var shapes = slide.getShapes();
    var tables = slide.getTables();
    
    var replacements = {
      '{{اسم_المعلم}}': data.teacherName || 'غير محدد',
      '{{موضوع_الاجتماع}}': meetingTopic,
      '{{رقم_الاجتماع}}': meetingNumber,
      '{{مقر_الاجتماع}}': data.meetingLocation || 'غير محدد',
      '{{التخصص}}': data.specialty || 'غير محدد',
      '{{الفئة_المستهدفة}}': data.targetAudience || 'غير محدد',
      '{{الهدف}}': data.objective || 'غير محدد',
      '{{رئيس_القسم}}': data.departmentHead || 'غير محدد',
      '{{خطة_العمل_1}}': data.workPlan?.plan1 || '',
      '{{خطة_العمل_2}}': data.workPlan?.plan2 || '',
      '{{خطة_العمل_3}}': data.workPlan?.plan3 || '',
      '{{خطة_العمل_4}}': data.workPlan?.plan4 || '',
      '{{المخرجات_1}}': data.outputs?.output1 || '',
      '{{المخرجات_2}}': data.outputs?.output2 || '',
      '{{المخرجات_3}}': data.outputs?.output3 || '',
      '{{المخرجات_4}}': data.outputs?.output4 || '',
      '{{مدة_التنفيذ_1}}': data.durations?.duration1 || '',
      '{{مدة_تنفيذ_1}}': data.durations?.duration1 || '',
      '{{مدة_التنفيذ_2}}': data.durations?.duration2 || '',
      '{{مدة_تنفيذ_2}}': data.durations?.duration2 || '',
      '{{مدة_التنفيذ_3}}': data.durations?.duration3 || '',
      '{{مدة_تنفيذ_3}}': data.durations?.duration3 || '',
      '{{مدة_التنفيذ_4}}': data.durations?.duration4 || '',
      '{{مدة_تنفيذ_4}}': data.durations?.duration4 || '',
      '{{الحاضرون}}': data.attendees || 'غير محدد',
      '{{الحاضرون_8}}': attendees8Text,
      '{{عدد_الحاضرون}}': attendeesCount.toString(),
      '{{وقت_الحفظ}}': currentTime,
      '{{وقت_الحفظ_45}}': futureTimeFormatted,
      '{{التاريخ}}': hijriDate,
      '{{اليوم}}': dayName
    };
    for (var i = 1; i <= 8; i++) { var key = '{{الحاضرون_' + i + '}}'; replacements[key] = (i <= attendeesList.length) ? attendeesList[i - 1] : ''; }
    
    shapes.forEach(function(shape) {
      if (shape.getText) {
        var text = shape.getText().asString();
        var newText = text;
        for (var key in replacements) { newText = newText.replace(new RegExp(key, 'g'), replacements[key]); }
        if (newText !== text) shape.getText().setText(newText);
      }
    });
    tables.forEach(function(table) {
      var numRows = table.getNumRows();
      var numCols = table.getNumColumns();
      for (var row = 0; row < numRows; row++) {
        for (var col = 0; col < numCols; col++) {
          var cell = table.getCell(row, col);
          var text = cell.getText().asString();
          var newText = text;
          for (var key in replacements) { newText = newText.replace(new RegExp(key, 'g'), replacements[key]); }
          if (newText !== text) cell.getText().setText(newText);
        }
      }
    });
  } catch(error) { throw error; }
}

// ====== دوال متابعة التنفيذ ======

function getActiveTopics() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    if (!dataSheet) return [];
    var hColumn = dataSheet.getRange('H:H').getValues();
    var iColumn = dataSheet.getRange('I:I').getValues();
    var jColumn = dataSheet.getRange('J:J').getValues();
    var now = new Date();
    now.setHours(0, 0, 0, 0);
    var activeTopics = [];
    for (var i = 0; i < hColumn.length; i++) {
      var topic = hColumn[i][0] ? hColumn[i][0].toString().trim() : '';
      if (topic) {
        var startDate = iColumn[i] && iColumn[i][0] ? new Date(iColumn[i][0]) : null;
        var endDate = jColumn[i] && jColumn[i][0] ? new Date(jColumn[i][0]) : null;
        var cleanTopic = topic.replace(/^\[\d+\]\s*/, '');
        if (!startDate || !endDate) { activeTopics.push(cleanTopic); }
        else if (now >= startDate && now <= endDate) { activeTopics.push(cleanTopic); }
      }
    }
    return activeTopics;
  } catch(e) { return []; }
}

// ====== جلب قائمة الرؤساء (من الصف 2) ======
function getLeadersList() {
  try {
    Logger.log('🔄 جلب قائمة الرؤساء...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var leadersSheet = sheet.getSheetByName('الرؤساء');
    
    if (!leadersSheet) {
      Logger.log('❌ شيت الرؤساء غير موجود');
      return [];
    }
    
    // جلب من العمود B بدءاً من الصف 2
    var data = leadersSheet.getRange('B2:B').getValues();
    var leaders = [];
    
    for (var i = 0; i < data.length; i++) {
      var name = data[i][0] ? data[i][0].toString().trim() : '';
      if (name) {
        leaders.push(name);
      }
    }
    
    Logger.log('✅ تم جلب ' + leaders.length + ' رئيس');
    if (leaders.length > 0) {
      Logger.log('👔 الرؤساء: ' + leaders.join(', '));
    }
    
    return leaders;
    
  } catch (error) {
    Logger.log('❌ خطأ في جلب الرؤساء: ' + error.toString());
    return [];
  }
}

// ====== جلب جميع الموضوعات مع أرقامها (بدون فلترة) ======
function getAllTopicsWithNumbers() {
  try {
    Logger.log('🔄 جلب جميع الموضوعات مع أرقامها...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    
    if (!dataSheet) {
      Logger.log('❌ شيت البيانات غير موجود');
      return [];
    }
    
    // جلب الأعمدة H (الموضوع) و G (الرقم)
    var hColumn = dataSheet.getRange('H:H').getValues();
    var gColumn = dataSheet.getRange('G:G').getValues();
    
    var allTopics = [];
    var seenNumbers = {};
    
    for (var i = 0; i < hColumn.length; i++) {
      var topic = hColumn[i][0] ? hColumn[i][0].toString().trim() : '';
      var number = gColumn[i][0] ? gColumn[i][0].toString().trim() : '';
      
      // ====== IMPORTANT: تجاهل الصف الأول إذا كان 'رقم الاجتماع' ======
      // التحقق من أن الرقم ليس 'رقم الاجتماع' أو أي نص غير رقمي
      var isNumeric = number && !isNaN(parseInt(number)) && parseInt(number).toString() === number;
      
      // تضمين جميع الموضوعات التي لها رقم رقمي وليست 'رقم الاجتماع'
      if (topic && number && isNumeric && !seenNumbers[number]) {
        seenNumbers[number] = true;
        var cleanTopic = topic.replace(/^\[\d+\]\s*/, '');
        allTopics.push({
          number: number,
          name: cleanTopic,
          full: topic
        });
      }
    }
    
    // ترتيب حسب الرقم
    allTopics.sort(function(a, b) {
      return parseInt(a.number) - parseInt(b.number);
    });
    
    Logger.log('✅ تم جلب ' + allTopics.length + ' موضوع');
    if (allTopics.length > 0) {
      Logger.log('📋 الأرقام: ' + allTopics.map(function(t) { return t.number; }).join(', '));
    }
    
    return allTopics;
    
  } catch (error) {
    Logger.log('❌ خطأ في جلب الموضوعات: ' + error.toString());
    return [];
  }
}


// ====== إصلاح حالة التنفيذ بناءً على سجل التنفيذ ======
function fixExecutionStatus() {
  try {
    Logger.log('🔄 جاري إصلاح حالات التنفيذ...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ====== 1. جلب سجل التنفيذ ======
    var logSheet = sheet.getSheetByName('تنفيذ_الرؤساء');
    if (!logSheet) {
      return { success: false, message: 'شيت تنفيذ_الرؤساء غير موجود' };
    }
    
    var logData = logSheet.getDataRange().getValues();
    if (logData.length <= 1) {
      return { success: false, message: 'لا يوجد سجل تنفيذ' };
    }
    
    // ====== 2. جلب شيت المتابعة ======
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    if (!trackerSheet) {
      return { success: false, message: 'شيت متابعة_التنفيذ غير موجود' };
    }
    
    var trackerData = trackerSheet.getDataRange().getValues();
    var headers = trackerData[0] || [];
    
    // ====== 3. بناء خريطة للبحث السريع ======
    var executionMap = {};
    
    // قراءة سجل التنفيذ (من الصف 2)
    for (var i = 1; i < logData.length; i++) {
      var row = logData[i];
      if (!row || row.length < 4) continue;
      
      var topicName = row[1] ? row[1].toString().trim() : '';
      var topicNumber = row[2] ? row[2].toString().trim() : '';
      var leaderName = row[3] ? row[3].toString().trim() : '';
      var status = row[4] ? row[4].toString().trim() : '';
      
      if (!leaderName || !topicNumber) continue;
      if (status !== '✅ تم التنفيذ') continue;
      
      var key = leaderName + '|' + topicNumber;
      executionMap[key] = true;
      Logger.log('📋 سجل: ' + leaderName + ' -> الموضوع ' + topicNumber);
    }
    
    Logger.log('✅ عدد سجلات التنفيذ: ' + Object.keys(executionMap).length);
    
    // ====== 4. تحديث شيت المتابعة ======
    var updated = 0;
    
    // البحث عن أعمدة الموضوعات (تجاهل '#' و 'رئيس القسم' و 'المجموع')
    var topicColumns = {};
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j] ? headers[j].toString().trim() : '';
      
      // تخطي الأعمدة غير الرقمية
      var match = header.match(/^\[(\d+)\]/);
      if (match) {
        var num = match[1];
        // التأكد من أن الرقم رقمي وليس 'رقم الاجتماع'
        if (!isNaN(parseInt(num)) && parseInt(num).toString() === num) {
          topicColumns[num] = j + 1; // العمود (1-indexed)
        }
      }
    }
    
    Logger.log('📊 أعمدة الموضوعات: ' + Object.keys(topicColumns).length);
    Logger.log('📊 الأرقام: ' + Object.keys(topicColumns).join(', '));
    
    // تحديث كل صف
    for (var rowIdx = 1; rowIdx < trackerData.length; rowIdx++) {
      var row = trackerData[rowIdx];
      if (!row || row.length < 2) continue;
      
      var leaderName = row[1] ? row[1].toString().trim() : '';
      if (!leaderName) continue;
      
      // التحقق من كل موضوع
      var execCount = 0;
      var totalTopics = Object.keys(topicColumns).length;
      
      for (var topicNum in topicColumns) {
        var col = topicColumns[topicNum];
        var key = leaderName + '|' + topicNum;
        
        // التحقق مما إذا كان هذا الرئيس قد نفذ هذا الموضوع
        if (executionMap[key]) {
          // تحديث الخلية إلى ✅
          trackerSheet.getRange(rowIdx + 1, col).setValue('✅');
          execCount++;
          updated++;
          Logger.log('✅ تحديث: ' + leaderName + ' -> الموضوع ' + topicNum);
        } else {
          // إذا كانت الخلية فارغة، ضع ☐
          var currentValue = trackerSheet.getRange(rowIdx + 1, col).getValue();
          if (!currentValue || currentValue === '') {
            trackerSheet.getRange(rowIdx + 1, col).setValue('☐');
          }
        }
      }
      
      // تحديث المجموع
      var lastCol = headers.length;
      trackerSheet.getRange(rowIdx + 1, lastCol).setValue(execCount + '/' + totalTopics);
      
      // تحديث وقت التنفيذ إذا كان هناك تنفيذ
      if (execCount > 0) {
        var now = new Date();
        var timeStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm');
        trackerSheet.getRange(rowIdx + 1, 1).setValue(timeStr);
      }
    }
    
    Logger.log('✅ تم تحديث ' + updated + ' خلية');
    
    return {
      success: true,
      message: 'تم تحديث ' + updated + ' حالة تنفيذ',
      updated: updated,
      totalExecutions: Object.keys(executionMap).length,
      totalTopics: Object.keys(topicColumns).length
    };
    
  } catch (error) {
    Logger.log('❌ خطأ: ' + error.toString());
    return { success: false, message: error.message };
  }
}

function getLeadersExecutionData() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    
    if (!trackerSheet) {
      return {
        success: true,
        headers: ['#', 'رئيس القسم'],
        data: [],
        stats: { totalLeaders: 0, totalTopics: 0, executed: 0, notExecuted: 0 },
        topicNames: {}
      };
    }
    
    var data = trackerSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        headers: data[0] || ['#', 'رئيس القسم'],
        data: [],
        stats: { totalLeaders: 0, totalTopics: 0, executed: 0, notExecuted: 0 },
        topicNames: {}
      };
    }
    
    var headers = data[0] || [];
    var rows = data.slice(1);
    
    // ====== جلب أسماء الموضوعات من شيت البيانات ======
    var topicNames = {};
    var dataSheet = sheet.getSheetByName('البيانات');
    if (dataSheet) {
      var gColumn = dataSheet.getRange('G:G').getValues();
      var hColumn = dataSheet.getRange('H:H').getValues();
      for (var i = 0; i < gColumn.length; i++) {
        var num = gColumn[i][0] ? gColumn[i][0].toString().trim() : '';
        var topic = hColumn[i][0] ? hColumn[i][0].toString().trim() : '';
        if (num && topic) {
          topicNames[num] = topic;
        }
      }
    }
    
    var stats = {
      totalLeaders: rows.length,
      totalTopics: Math.max(0, headers.length - 3),
      executed: 0,
      notExecuted: 0
    };
    
    var leadersData = [];
    
    rows.forEach(function(row) {
      var leader = row[1] || '';
      if (!leader) return;
      
      var executions = {};
      var executedCount = 0;
      
      for (var i = 2; i < row.length - 1 && i < headers.length - 1; i++) {
        var topicHeader = headers[i] || '';
        var status = row[i] || '☐';
        executions[topicHeader] = status;
        if (status === '✅') {
          executedCount++;
          stats.executed++;
        } else {
          stats.notExecuted++;
        }
      }
      
      leadersData.push({
        name: leader,
        executions: executions,
        executedCount: executedCount,
        totalTopics: stats.totalTopics
      });
    });
    
    return {
      success: true,
      headers: headers,
      data: leadersData,
      stats: stats,
      topicNames: topicNames
    };
    
  } catch (error) {
    Logger.log('❌ خطأ في جلب البيانات: ' + error.toString());
    return {
      success: false,
      message: error.message,
      headers: ['#', 'رئيس القسم'],
      data: [],
      stats: { totalLeaders: 0, totalTopics: 0, executed: 0, notExecuted: 0 },
      topicNames: {}
    };
  }
}
// ====== تبديل حالة التنفيذ (مع تحديث السجل بدلاً من الإضافة) ======
function toggleLeaderExecution(leaderName, topicNumber) {
  try {
    Logger.log('🔄 تبديل حالة التنفيذ لـ ' + leaderName + ' - الموضوع رقم ' + topicNumber);
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    
    if (!trackerSheet) {
      return { success: false, message: 'شيت متابعة_التنفيذ غير موجود' };
    }
    
    var data = trackerSheet.getDataRange().getValues();
    var headers = data[0] || [];
    var leaderRow = -1;
    var topicCol = -1;
    
    // البحث عن الرئيس
    for (var i = 1; i < data.length; i++) {
      if (data[i] && data[i][1] && data[i][1].toString().trim() === leaderName) {
        leaderRow = i + 1;
        break;
      }
    }
    
    if (leaderRow === -1) {
      return { success: false, message: 'الرئيس غير موجود: ' + leaderName };
    }
    
    // البحث عن الموضوع باستخدام الرقم
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j] ? headers[j].toString().trim() : '';
      if (header === '[' + topicNumber + ']') {
        topicCol = j + 1;
        break;
      }
    }
    
    if (topicCol === -1) {
      return { success: false, message: 'الموضوع غير موجود: ' + topicNumber };
    }
    
    // ====== قراءة الحالة الحالية ======
    var currentStatus = trackerSheet.getRange(leaderRow, topicCol).getValue();
    var newStatus = (currentStatus === '✅') ? '☐' : '✅';
    
    // ====== تحديث شيت المتابعة ======
    trackerSheet.getRange(leaderRow, topicCol).setValue(newStatus);
    
    // تحديث المجموع
    var lastCol = headers.length;
    var totalTopics = headers.length - 3;
    var execCount = 0;
    for (var k = 2; k < headers.length - 1; k++) {
      var status = trackerSheet.getRange(leaderRow, k + 1).getValue();
      if (status === '✅') execCount++;
    }
    trackerSheet.getRange(leaderRow, lastCol).setValue(execCount + '/' + totalTopics);
    
    // تحديث وقت التنفيذ
    var now = new Date();
    var timeStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm');
    if (newStatus === '✅') {
      trackerSheet.getRange(leaderRow, 1).setValue(timeStr);
    } else {
      // إذا تم الإلغاء، امسح وقت التنفيذ
      trackerSheet.getRange(leaderRow, 1).setValue('');
    }
    
    // ====== تحديث شيت تنفيذ_الرؤساء (تحديث بدلاً من إضافة) ======
    var logSheet = sheet.getSheetByName('تنفيذ_الرؤساء');
    if (!logSheet) {
      logSheet = sheet.insertSheet('تنفيذ_الرؤساء');
      logSheet.getRange(1, 1, 1, 6).setValues([['التاريخ', 'الموضوع', 'رقم الموضوع', 'اسم الرئيس', 'حالة التنفيذ', 'وقت التنفيذ']]);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      logSheet.getRange(1, 1, 1, 6).setBackground('#4f6ef7');
      logSheet.getRange(1, 1, 1, 6).setFontColor('white');
    }
    
    // ====== جلب اسم الموضوع ======
    var topicName = '';
    var dataSheet = sheet.getSheetByName('البيانات');
    if (dataSheet) {
      var gColumn = dataSheet.getRange('G:G').getValues();
      var hColumn = dataSheet.getRange('H:H').getValues();
      for (var m = 0; m < gColumn.length; m++) {
        var num = gColumn[m][0] ? gColumn[m][0].toString().trim() : '';
        if (num === topicNumber) {
          topicName = hColumn[m][0] ? hColumn[m][0].toString().trim() : '';
          break;
        }
      }
    }
    
    // ====== البحث عن سجل موجود ======
    var logData = logSheet.getDataRange().getValues();
    var existingRow = -1;
    
    // البحث عن سجل بنفس الرئيس والموضوع
    for (var r = 1; r < logData.length; r++) {
      var row = logData[r];
      if (!row || row.length < 4) continue;
      var logLeader = row[3] ? row[3].toString().trim() : '';
      var logTopic = row[2] ? row[2].toString().trim() : '';
      if (logLeader === leaderName && logTopic === topicNumber) {
        existingRow = r + 1;
        break;
      }
    }
    
    var dateStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm:ss');
    
    if (newStatus === '✅') {
      // ====== حالة: تم التنفيذ ======
      if (existingRow === -1) {
        // لا يوجد سجل، أضف جديد
        var lastRow = logSheet.getLastRow() + 1;
        logSheet.getRange(lastRow, 1, 1, 6).setValues([
          [dateStr, topicName || 'الموضوع ' + topicNumber, topicNumber, leaderName, '✅ تم التنفيذ', dateStr]
        ]);
        Logger.log('✅ تم إضافة سجل جديد');
      } else {
        // يوجد سجل، قم بتحديثه
        logSheet.getRange(existingRow, 1).setValue(dateStr); // تحديث التاريخ
        logSheet.getRange(existingRow, 5).setValue('✅ تم التنفيذ'); // تحديث الحالة
        logSheet.getRange(existingRow, 6).setValue(dateStr); // تحديث وقت التنفيذ
        Logger.log('✅ تم تحديث السجل الموجود');
      }
    } else {
      // ====== حالة: تم الإلغاء ======
      if (existingRow !== -1) {
        // حذف السجل
        logSheet.deleteRow(existingRow);
        Logger.log('🗑️ تم حذف السجل');
      }
      // إذا لم يوجد سجل، لا تفعل شيئاً
    }
    
    return {
      success: true,
      message: newStatus === '✅' ? 'تم التنفيذ بنجاح' : 'تم إلغاء التنفيذ',
      leader: leaderName,
      topic: topicNumber,
      status: newStatus
    };
    
  } catch (error) {
    Logger.log('❌ خطأ في تبديل الحالة: ' + error.toString());
    return { success: false, message: error.message };
  }
}
// ====== إعادة ترقيم الموضوعات تلقائياً ======
// ====== إعادة ترقيم الموضوعات تلقائياً ======
function renumberTopics() {
  try {
    Logger.log('🔄 جاري إعادة ترقيم الموضوعات...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    
    if (!dataSheet) {
      return { success: false, message: 'شيت البيانات غير موجود' };
    }
    
    // جلب جميع البيانات من الأعمدة G و H
    var gColumn = dataSheet.getRange('G:G').getValues();
    var hColumn = dataSheet.getRange('H:H').getValues();
    
    // جمع الموضوعات مع أرقامها
    var topics = [];
    for (var i = 0; i < gColumn.length; i++) {
      var number = gColumn[i][0] ? gColumn[i][0].toString().trim() : '';
      var topic = hColumn[i][0] ? hColumn[i][0].toString().trim() : '';
      
      // التحقق من أن الرقم رقمي وليس 'رقم الاجتماع'
      var isNumeric = number && !isNaN(parseInt(number)) && parseInt(number).toString() === number;
      
      if (topic && isNumeric) {
        topics.push({
          row: i + 1,
          number: parseInt(number),
          topic: topic
        });
      }
    }
    
    if (topics.length === 0) {
      return { success: false, message: 'لا توجد موضوعات لإعادة الترقيم' };
    }
    
    // ترتيب حسب الصف (لضمان الترتيب الصحيح)
    topics.sort(function(a, b) {
      return a.row - b.row;
    });
    
    // إعادة الترقيم من 1 حسب ترتيب الصفوف
    var newNumber = 1;
    var updated = 0;
    
    topics.forEach(function(item) {
      // تحديث الرقم
      dataSheet.getRange(item.row, 7).setValue(newNumber); // العمود G
      if (item.number !== newNumber) {
        updated++;
        Logger.log('✅ تم تحديث الموضوع "' + item.topic + '" من ' + item.number + ' إلى ' + newNumber);
      }
      newNumber++;
    });
    
    Logger.log('✅ تم إعادة ترقيم ' + updated + ' موضوع');
    
    return {
      success: true,
      message: 'تم إعادة ترقيم ' + updated + ' موضوع',
      updated: updated,
      total: topics.length
    };
    
  } catch (error) {
    Logger.log('❌ خطأ في إعادة الترقيم: ' + error.toString());
    return { success: false, message: error.message };
  }
}
// ====== تحديث حالة التنفيذ عند حفظ PDF (معدل) ======
function updateLeaderExecution(leaderName, topicName) {
  try {
    Logger.log('🔄 تحديث التنفيذ لـ ' + leaderName + ' - ' + topicName);
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ====== 1. البحث عن رقم الموضوع ======
    var dataSheet = sheet.getSheetByName('البيانات');
    var topicNumber = '';
    
    if (dataSheet) {
      var hColumn = dataSheet.getRange('H:H').getValues();
      var gColumn = dataSheet.getRange('G:G').getValues();
      
      for (var i = 0; i < hColumn.length; i++) {
        var topic = hColumn[i][0] ? hColumn[i][0].toString().trim() : '';
        var number = gColumn[i][0] ? gColumn[i][0].toString().trim() : '';
        
        // مقارنة باسم الموضوع (بدون أرقام)
        var cleanTopic = topic.replace(/^\[\d+\]\s*/, '');
        var cleanName = topicName.replace(/^\[\d+\]\s*/, '');
        
        if (cleanTopic === cleanName && number) {
          topicNumber = number;
          Logger.log('✅ تم العثور على رقم الموضوع: ' + topicNumber);
          break;
        }
      }
    }
    
    // إذا لم يتم العثور على الرقم، استخدم الرقم من اسم الموضوع
    if (!topicNumber) {
      var match = topicName.match(/^\[(\d+)\]/);
      if (match) {
        topicNumber = match[1];
        Logger.log('✅ استخراج الرقم من اسم الموضوع: ' + topicNumber);
      }
    }
    
    if (!topicNumber) {
      Logger.log('⚠️ لم يتم العثور على رقم الموضوع: ' + topicName);
      return;
    }
    
    // ====== 2. تحديث شيت متابعة_التنفيذ ======
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    
    if (trackerSheet) {
      var data = trackerSheet.getDataRange().getValues();
      var leaderRow = -1;
      var topicCol = -1;
      var headers = data[0] || [];
      
      // البحث عن الرئيس
      for (var i = 1; i < data.length; i++) {
        if (data[i] && data[i][1] && data[i][1].toString().trim() === leaderName) {
          leaderRow = i + 1;
          break;
        }
      }
      
      // البحث عن عمود الموضوع باستخدام الرقم بين قوسين
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j] ? headers[j].toString().trim() : '';
        // البحث عن '[' + رقم + ']'
        if (header === '[' + topicNumber + ']') {
          topicCol = j + 1;
          break;
        }
      }
      
      // إذا لم يتم العثور بالرقم، حاول البحث بالنص
      if (topicCol === -1) {
        for (var k = 0; k < headers.length; k++) {
          var headerText = headers[k] ? headers[k].toString().trim() : '';
          if (headerText === topicName || headerText === topicNumber) {
            topicCol = k + 1;
            break;
          }
        }
      }
      
      if (leaderRow !== -1 && topicCol !== -1) {
        // تحديث الحالة إلى ✅
        trackerSheet.getRange(leaderRow, topicCol).setValue('✅');
        Logger.log('✅ تم تحديث الحالة في شيت المتابعة للصف ' + leaderRow + ' العمود ' + topicCol);
        
        // تحديث المجموع
        var lastCol = headers.length;
        var totalTopics = headers.length - 3;
        var execCount = 0;
        for (var x = 2; x < headers.length - 1; x++) {
          var status = trackerSheet.getRange(leaderRow, x + 1).getValue();
          if (status === '✅') execCount++;
        }
        trackerSheet.getRange(leaderRow, lastCol).setValue(execCount + '/' + totalTopics);
        Logger.log('✅ تم تحديث المجموع: ' + execCount + '/' + totalTopics);
        
        // تسجيل الوقت
        var now = new Date();
        var timeStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm');
        trackerSheet.getRange(leaderRow, 1).setValue(timeStr);
        
      } else {
        if (leaderRow === -1) Logger.log('⚠️ الرئيس غير موجود في شيت المتابعة: ' + leaderName);
        if (topicCol === -1) Logger.log('⚠️ الموضوع غير موجود في شيت المتابعة: ' + topicNumber);
      }
    } else {
      Logger.log('⚠️ شيت متابعة_التنفيذ غير موجود');
    }
    
    // ====== 3. تسجيل في شيت تنفيذ_الرؤساء ======
    var logSheet = sheet.getSheetByName('تنفيذ_الرؤساء');
    if (!logSheet) {
      logSheet = sheet.insertSheet('تنفيذ_الرؤساء');
      logSheet.getRange(1, 1, 1, 6).setValues([['التاريخ', 'الموضوع', 'رقم الموضوع', 'اسم الرئيس', 'الحالة', 'وقت التنفيذ']]);
      logSheet.getRange(1, 1, 1, 6).setFontWeight('bold');
      logSheet.getRange(1, 1, 1, 6).setBackground('#4f6ef7');
      logSheet.getRange(1, 1, 1, 6).setFontColor('white');
    }
    
    var now = new Date();
    var dateStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm:ss');
    var lastRow = logSheet.getLastRow() + 1;
    logSheet.getRange(lastRow, 1, 1, 6).setValues([
      [dateStr, topicName, topicNumber, leaderName, '✅ تم التنفيذ', dateStr]
    ]);
    
    Logger.log('✅ تم تسجيل التنفيذ في شيت السجل');
    
  } catch (error) {
    Logger.log('❌ خطأ في updateLeaderExecution: ' + error.toString());
  }
}

// ====== إضافة رئيس تخصص جديد (مع جلب الكود) ======
function addSpecialtyLeader(leaderData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var leadersSheet = sheet.getSheetByName('الرؤساء');
    
    if (!leadersSheet) throw new Error('شيت الرؤساء غير موجود');
    
    // التحقق من عدم وجود الاسم مسبقاً
    var existingData = leadersSheet.getRange('B:B').getValues();
    for (var i = 0; i < existingData.length; i++) {
      if (existingData[i][0] && existingData[i][0].toString().trim() === leaderData.name) {
        throw new Error('الاسم موجود بالفعل');
      }
    }
    
    // جلب الكود من شيت الحاضرون
    var code = leaderData.code || '0000';
    
    // إذا لم يتم إرسال الكود، حاول جلبها من شيت الحاضرون
    if (!code || code === '0000' || code === '0') {
      var attendeesSheet = sheet.getSheetByName('الحاضرون');
      if (attendeesSheet) {
        var attendeeData = attendeesSheet.getRange('B2:C').getValues();
        for (var i = 0; i < attendeeData.length; i++) {
          var name = attendeeData[i][0] ? attendeeData[i][0].toString().trim() : '';
          var attendeeCode = attendeeData[i][1] ? attendeeData[i][1].toString().trim() : '';
          if (name === leaderData.name && attendeeCode) {
            code = attendeeCode;
            break;
          }
        }
      }
    }
    
    // إذا لم يتم العثور على كود، استخدم الرقم التسلسلي
    if (!code || code === '0' || code === '0000') {
      var lastRow = leadersSheet.getLastRow();
      code = (lastRow + 1).toString().padStart(4, '0');
    }
    
    var lastRow = leadersSheet.getLastRow();
    var newRow = lastRow + 1;
    
    leadersSheet.getRange('B' + newRow).setValue(leaderData.name);
    leadersSheet.getRange('C' + newRow).setValue(code);
    
    // تحديث شيت المتابعة
    try {
      resetTrackerSheet();
    } catch (e) {
      // تجاهل
    }
    
    return {
      success: true,
      message: 'تم إضافة رئيس التخصص بنجاح',
      name: leaderData.name,
      code: code
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

// ====== حذف رئيس من الشيت ======
function deleteLeaderFromSheet(leaderName) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var leadersSheet = sheet.getSheetByName('الرؤساء');
    
    if (!leadersSheet) throw new Error('شيت الرؤساء غير موجود');
    
    var data = leadersSheet.getRange('B:B').getValues();
    var rowToDelete = -1;
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === leaderName) {
        rowToDelete = i + 1;
        break;
      }
    }
    
    if (rowToDelete === -1) {
      throw new Error('الرئيس غير موجود');
    }
    
    leadersSheet.deleteRow(rowToDelete);
    
    // تحديث شيت المتابعة
    try {
      resetTrackerSheet();
    } catch(e) {}
    
    return {
      success: true,
      message: 'تم حذف رئيس التخصص بنجاح',
      name: leaderName
    };
    
  } catch (error) {
    return {
      success: false,
      message: error.message
    };
  }
}

// ====== تسجيل في سجل التنفيذ ======
function logExecution(leaderName, topicName) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSheet = sheet.getSheetByName('تنفيذ_الرؤساء');
    
    if (!logSheet) {
      logSheet = sheet.insertSheet('تنفيذ_الرؤساء');
      logSheet.getRange(1, 1, 1, 5).setValues([['التاريخ', 'الموضوع', 'اسم الرئيس', 'الحالة', 'وقت التنفيذ']]);
      logSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
      logSheet.getRange(1, 1, 1, 5).setBackground('#4f6ef7');
      logSheet.getRange(1, 1, 1, 5).setFontColor('white');
    }
    
    var now = new Date();
    var dateStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm:ss');
    var lastRow = logSheet.getLastRow() + 1;
    logSheet.getRange(lastRow, 1, 1, 5).setValues([[dateStr, topicName, leaderName, '✅ تم التنفيذ', dateStr]]);
    
  } catch (error) {
    Logger.log('⚠️ خطأ في تسجيل التنفيذ: ' + error.toString());
  }
}

function refreshAttendees() {
  try {
    var teachers = fetchTeachersData();
    var projectSheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var attendeesSheet = projectSheet.getSheetByName('الحاضرون');
    if (!attendeesSheet) throw new Error('شيت الحاضرون غير موجود');
    var lastRow = attendeesSheet.getLastRow();
    if (lastRow > 1) attendeesSheet.getRange(2, 2, lastRow - 1, 5).clearContent();
    var dataToWrite = [];
    teachers.forEach(function(teacher) {
      dataToWrite.push([teacher.name, teacher.code, teacher.department, teacher.specialty, teacher.phone]);
    });
    if (dataToWrite.length > 0) attendeesSheet.getRange(2, 2, dataToWrite.length, 5).setValues(dataToWrite);
    return { success: true, message: 'تم تحديث بيانات الحاضرون', count: dataToWrite.length };
  } catch(error) {
    return { success: false, message: error.message };
  }
}

// function fetchTeachersData() {
//   try {
//     var teachersSheet = SpreadsheetApp.openById(TEACHERS_SPREADSHEET_ID);
//     var sheet = teachersSheet.getSheetByName('المعلمون');
//     if (!sheet) throw new Error('شيت المعلمون غير موجود');
//     var data = sheet.getDataRange().getValues();
//     var teachers = [];
//     for (var i = 1; i < data.length; i++) {
//       var row = data[i];
//       if (row[0] && row[0].toString().trim() !== '') {
//         teachers.push({
//           name: row[0] ? row[0].toString().trim() : '',
//           code: row[1] ? row[1].toString().trim() : '',
//           department: row[2] ? row[2].toString().trim() : '',
//           specialty: row[3] ? row[3].toString().trim() : '',
//           phone: row[4] ? row[4].toString().trim() : ''
//         });
//       }
//     }
//     return teachers;
//   } catch(error) {
//     throw new Error('فشل في جلب بيانات المعلمين: ' + error.message);
//   }
// }

function checkTriggerPermissions() {
  try { var triggers = ScriptApp.getProjectTriggers(); return { success: true, hasTriggers: triggers.length > 0 }; } 
  catch(error) { return { success: false, message: error.toString(), requiredScope: 'https://www.googleapis.com/auth/script.scriptapp' }; }
}

function toggleAutoUpdate(enable) {
  try {
    var permissionCheck = checkTriggerPermissions();
    if (!permissionCheck.success) {
      return { success: false, message: 'لا يمكن تغيير حالة المؤقت بسبب نقص الصلاحيات', triggerEnabled: false, permissionError: true };
    }
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'updateAttendeesSheet') ScriptApp.deleteTrigger(trigger);
    });
    if (enable) {
      ScriptApp.newTrigger('updateAttendeesSheet').timeBased().everyHours(1).create();
      return { success: true, message: 'تم تفعيل التحديث التلقائي كل ساعة', triggerEnabled: true };
    } else {
      return { success: true, message: 'تم إيقاف التحديث التلقائي', triggerEnabled: false };
    }
  } catch(error) {
    return { success: false, message: error.message, triggerEnabled: false };
  }
}

function getAttendeesUpdateStatus() {
  try {
    var permissionCheck = checkTriggerPermissions();
    if (!permissionCheck.success) {
      return { success: true, hasTrigger: false, message: 'الصلاحيات غير كافية', permissionError: true };
    }
    var triggers = ScriptApp.getProjectTriggers();
    var hasTrigger = false;
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'updateAttendeesSheet') hasTrigger = true;
    });
    return { success: true, hasTrigger: hasTrigger, message: hasTrigger ? 'المؤقت مفعل (كل ساعة)' : 'لا يوجد مؤقت', permissionError: false };
  } catch(error) {
    return { success: false, hasTrigger: false, message: error.message, permissionError: true };
  }
}

function scheduleAttendeesUpdate() {
  try {
    var permissionCheck = checkTriggerPermissions();
    if (!permissionCheck.success) {
      var result = refreshAttendees();
      return { success: true, message: 'تم تحديث البيانات (بدون مؤقت)', result: result, triggerEnabled: false, permissionError: true };
    }
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'updateAttendeesSheet') ScriptApp.deleteTrigger(trigger);
    });
    ScriptApp.newTrigger('updateAttendeesSheet').timeBased().everyHours(1).create();
    var result = refreshAttendees();
    return { success: true, message: 'تم إنشاء المؤقت وتحديث البيانات', result: result, triggerEnabled: true, permissionError: false };
  } catch(error) {
    try { var result = refreshAttendees(); return { success: true, message: 'تم تحديث البيانات (تعذر إنشاء المؤقت)', result: result, triggerEnabled: false }; }
    catch(e) { throw new Error('فشل في تحديث البيانات: ' + e.message); }
  }
}

// ====== جلب الموضوعات النشطة مع أرقامها ======
function getActiveTopicsWithNumbers() {
  try {
    Logger.log('🔄 جلب الموضوعات النشطة مع أرقامها...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = sheet.getSheetByName('البيانات');
    if (!dataSheet) {
      Logger.log('❌ شيت البيانات غير موجود');
      return [];
    }
    
    // جلب البيانات من الأعمدة المطلوبة
    var hColumn = dataSheet.getRange('H:H').getValues(); // الموضوع
    var gColumn = dataSheet.getRange('G:G').getValues(); // الرقم
    var iColumn = dataSheet.getRange('I:I').getValues(); // تاريخ البدء
    var jColumn = dataSheet.getRange('J:J').getValues(); // تاريخ الانتهاء
    
    var now = new Date();
    now.setHours(0, 0, 0, 0);
    
    var activeTopics = [];
    
    for (var i = 0; i < hColumn.length; i++) {
      var topic = hColumn[i][0] ? hColumn[i][0].toString().trim() : '';
      var number = gColumn[i][0] ? gColumn[i][0].toString().trim() : '';
      
      if (topic && number) {
        var startDate = iColumn[i] && iColumn[i][0] ? new Date(iColumn[i][0]) : null;
        var endDate = jColumn[i] && jColumn[i][0] ? new Date(jColumn[i][0]) : null;
        var cleanTopic = topic.replace(/^\[\d+\]\s*/, '');
        
        // التحقق من أن الموضوع مفعل
        var isActive = false;
        if (!startDate || !endDate) {
          isActive = true; // إذا لم يكن هناك تواريخ، اعتبره مفعلاً
        } else if (now >= startDate && now <= endDate) {
          isActive = true;
        }
        
        if (isActive) {
          activeTopics.push({
            number: number,
            name: cleanTopic,
            full: topic
          });
        }
      }
    }
    
    Logger.log('✅ تم جلب ' + activeTopics.length + ' موضوع نشط');
    return activeTopics;
    
  } catch (error) {
    Logger.log('❌ خطأ في جلب الموضوعات: ' + error.toString());
    return [];
  }
}

// ====== إصلاح عناوين شيت المتابعة ======
function fixTrackerHeaders() {
  try {
    Logger.log('🔄 إصلاح عناوين شيت متابعة_التنفيذ...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    
    if (!trackerSheet) {
      return { success: false, message: 'شيت متابعة_التنفيذ غير موجود' };
    }
    
    // جلب جميع الموضوعات
    var allTopics = getAllTopicsWithNumbers();
    var leaders = getLeadersList();
    
    // بناء العناوين الجديدة
    var headers = ['#', 'رئيس القسم'];
    allTopics.forEach(function(t) {
      headers.push('[' + t.number + ']');
    });
    headers.push('المجموع');
    
    // تحديث العناوين
    trackerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    trackerSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    trackerSheet.getRange(1, 1, 1, headers.length).setBackground('#f0f4ff');
    trackerSheet.getRange(1, 1, 1, headers.length).setFontColor('#000000');
    trackerSheet.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
    
    // ضبط عرض الأعمدة
    trackerSheet.setColumnWidth(1, 50);
    trackerSheet.setColumnWidth(2, 200);
    for (var i = 3; i <= headers.length; i++) {
      trackerSheet.setColumnWidth(i, 60);
    }
    
    // إعادة تعبئة الرؤساء
    var lastRow = trackerSheet.getLastRow();
    if (lastRow > 1) {
      trackerSheet.deleteRows(2, lastRow - 1);
    }
    
    if (leaders.length > 0) {
      var rows = [];
      var serial = 1;
      leaders.forEach(function(leader) {
        var row = [serial, leader];
        for (var i = 0; i < allTopics.length; i++) {
          row.push('☐');
        }
        row.push('0/' + allTopics.length);
        rows.push(row);
        serial++;
      });
      trackerSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    trackerSheet.setFrozenRows(1);
    
    return {
      success: true,
      message: 'تم إصلاح العناوين بنجاح',
      headers: headers
    };
    
  } catch (error) {
    Logger.log('❌ خطأ في إصلاح العناوين: ' + error.toString());
    return { success: false, message: error.message };
  }
}

// ====== إعادة إنشاء شيت متابعة_التنفيذ ======
function resetTrackerSheet() {
  try {
    Logger.log('🔄 إعادة إنشاء شيت متابعة_التنفيذ...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    
    // حذف الشيت القديم إذا وجد
    if (trackerSheet) {
      sheet.deleteSheet(trackerSheet);
      Logger.log('✅ تم حذف الشيت القديم');
    }
    
    // إنشاء شيت جديد
    trackerSheet = sheet.insertSheet('متابعة_التنفيذ');
    
    // ====== جلب جميع الموضوعات (بدون فلترة) ======
    var allTopics = getAllTopicsWithNumbers();
    var leaders = getLeadersList();
    
    Logger.log('📚 عدد الموضوعات الكلي: ' + allTopics.length);
    Logger.log('👔 عدد الرؤساء: ' + leaders.length);
    
    // ====== بناء العناوين ======
    var headers = ['#', 'رئيس القسم'];
    allTopics.forEach(function(t) {
      headers.push('[' + t.number + ']');
    });
    headers.push('المجموع');
    
    // كتابة العناوين
    trackerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    trackerSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    trackerSheet.getRange(1, 1, 1, headers.length).setBackground('#f0f4ff');
    trackerSheet.getRange(1, 1, 1, headers.length).setFontColor('#000000');
    trackerSheet.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
    
    // ضبط عرض الأعمدة
    trackerSheet.setColumnWidth(1, 50);
    trackerSheet.setColumnWidth(2, 200);
    for (var i = 3; i <= headers.length; i++) {
      trackerSheet.setColumnWidth(i, 60);
    }
    
    // ====== إضافة الرؤساء من الصف 2 ======
    if (leaders.length > 0) {
      var rows = [];
      var serial = 1;
      leaders.forEach(function(leader) {
        var row = [serial, leader];
        for (var i = 0; i < allTopics.length; i++) {
          row.push('☐');
        }
        row.push('0/' + allTopics.length);
        rows.push(row);
        serial++;
      });
      
      trackerSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      Logger.log('✅ تم إضافة ' + rows.length + ' رئيس');
    }
    
    trackerSheet.setFrozenRows(1);
    
    Logger.log('✅ تم إنشاء شيت جديد بالعناوين: ' + headers.join(' | '));
    
    return {
      success: true,
      message: 'تم إعادة إنشاء شيت المتابعة',
      headers: headers,
      topicsCount: allTopics.length,
      leadersCount: leaders.length
    };
    
  } catch (error) {
    Logger.log('❌ خطأ في إعادة إنشاء الشيت: ' + error.toString());
    return { success: false, message: error.message };
  }
}

// ====== مزامنة البيانات (نسخة مُعاد كتابتها بالكامل) ======
function syncTrackerData() {
  try {
    Logger.log('🔄 بدء مزامنة بيانات المتابعة...');
    
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ====== 1. جلب جميع الموضوعات ======
    var allTopics = getAllTopicsWithNumbers();
    Logger.log('📚 عدد الموضوعات: ' + allTopics.length);
    
    if (allTopics.length === 0) {
      return { success: false, message: 'لا توجد موضوعات للمزامنة' };
    }
    
    // ====== 2. جلب جميع الرؤساء ======
    var leaders = getLeadersList();
    Logger.log('👔 عدد الرؤساء: ' + leaders.length);
    
    if (leaders.length === 0) {
      return { success: false, message: 'لا يوجد رؤساء للمزامنة' };
    }
    
    // ====== 3. جلب سجل التنفيذ ======
    var executionMap = {};
    var logSheet = sheet.getSheetByName('تنفيذ_الرؤساء');
    if (logSheet) {
      var logData = logSheet.getDataRange().getValues();
      for (var i = 1; i < logData.length; i++) {
        var row = logData[i];
        if (!row || row.length < 4) continue;
        
        var topicNumber = row[2] ? row[2].toString().trim() : '';
        var leaderName = row[3] ? row[3].toString().trim() : '';
        var status = row[4] ? row[4].toString().trim() : '';
        
        if (leaderName && topicNumber && status === '✅ تم التنفيذ') {
          var key = leaderName + '|' + topicNumber;
          executionMap[key] = true;
        }
      }
      Logger.log('📋 عدد سجلات التنفيذ: ' + Object.keys(executionMap).length);
    }
    
    // ====== 4. حذف الشيت القديم ======
    var trackerSheet = sheet.getSheetByName('متابعة_التنفيذ');
    if (trackerSheet) {
      sheet.deleteSheet(trackerSheet);
      Logger.log('✅ تم حذف الشيت القديم');
    }
    
    // ====== 5. إنشاء شيت جديد ======
    trackerSheet = sheet.insertSheet('متابعة_التنفيذ');
    Logger.log('✅ تم إنشاء شيت جديد');
    
    // ====== 6. بناء العناوين بشكل صحيح ======
    // العمود 1: # (رقم تسلسلي)
    // العمود 2: رئيس القسم
    // الأعمدة من 3 إلى نهاية: [رقم] لكل موضوع
    // العمود الأخير: المجموع
    
    var headers = ['#', 'رئيس القسم'];
    allTopics.forEach(function(t) {
      headers.push('[' + t.number + ']');
    });
    headers.push('المجموع');
    
    Logger.log('📊 العناوين: ' + headers.join(' | '));
    Logger.log('📊 عدد الأعمدة: ' + headers.length);
    
    // ====== 7. كتابة العناوين ======
    trackerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    trackerSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    trackerSheet.getRange(1, 1, 1, headers.length).setBackground('#f0f4ff');
    trackerSheet.getRange(1, 1, 1, headers.length).setFontColor('#000000');
    trackerSheet.getRange(1, 1, 1, headers.length).setHorizontalAlignment('center');
    
    // ضبط عرض الأعمدة
    trackerSheet.setColumnWidth(1, 50);
    trackerSheet.setColumnWidth(2, 200);
    for (var i = 3; i <= headers.length; i++) {
      trackerSheet.setColumnWidth(i, 60);
    }
    
    // ====== 8. إضافة الرؤساء مع حالات التنفيذ ======
    var rows = [];
    var serial = 1;
    var totalTopics = allTopics.length;
    
    leaders.forEach(function(leader) {
      var row = [serial, leader];
      var execCount = 0;
      
      // لكل موضوع، تحقق من حالة التنفيذ
      allTopics.forEach(function(topic) {
        var key = leader + '|' + topic.number;
        if (executionMap[key]) {
          row.push('✅');
          execCount++;
        } else {
          row.push('☐');
        }
      });
      
      // أضف المجموع
      row.push(execCount + '/' + totalTopics);
      rows.push(row);
      serial++;
    });
    
    if (rows.length > 0) {
      trackerSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      Logger.log('✅ تم إضافة ' + rows.length + ' رئيس');
    }
    
    trackerSheet.setFrozenRows(1);
    
    // ====== 9. تحديث وقت التنفيذ للرؤساء الذين نفذوا ======
    // البحث عن الرؤساء الذين لديهم تنفيذ وتحديث وقتهم
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];
      var leaderName = row[1];
      var hasExecution = false;
      
      // تحقق من وجود أي ✅ في الصف
      for (var c = 2; c < row.length - 1; c++) {
        if (row[c] === '✅') {
          hasExecution = true;
          break;
        }
      }
      
      if (hasExecution) {
        var now = new Date();
        var timeStr = Utilities.formatDate(now, 'Asia/Riyadh', 'dd/MM/yyyy HH:mm');
        trackerSheet.getRange(r + 2, 1).setValue(timeStr);
      }
    }
    
    Logger.log('✅ تمت المزامنة بنجاح');
    
    return {
      success: true,
      message: 'تمت المزامنة بنجاح',
      topicsCount: totalTopics,
      leadersCount: leaders.length,
      headers: headers,
      executionsCount: Object.keys(executionMap).length
    };
    
  } catch (error) {
    Logger.log('❌ خطأ في المزامنة: ' + error.toString());
    return { success: false, message: error.message };
  }
}
