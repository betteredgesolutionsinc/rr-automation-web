/** BESI Shift Assigner: writes planned hours to dated TK Time In/Out cells. */
const TK_SHIFT = { spreadsheetId: '', optionsTab: 'SHIFT_OPTIONS' };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('BESI Shift Assigner');
}

function setupShiftSystem() {
  const book = tkSpreadsheet_();
  let sheet = book.getSheetByName(TK_SHIFT.optionsTab);
  if (!sheet) sheet = book.insertSheet(TK_SHIFT.optionsTab);
  if (sheet.getLastRow() === 0) {
    sheet.getRange('A1:E1').setValues([['Shift ID','Shift Name','Start Time','End Time','Status']]);
    sheet.getRange('A1:E1').setBackground('#cfe9f7').setFontWeight('bold');
    sheet.getRange('C2:D5').setNumberFormat('@');
    sheet.getRange('A2:E5').setValues([
      ['NIGHT_21','Night Shift','21:00','06:00','ACTIVE'],
      ['MIDNIGHT_00','Midnight Shift','00:00','09:00','ACTIVE'],
      ['EARLY_01','Early Shift','01:00','10:00','ACTIVE'],
      ['MORNING_04','Morning Shift','04:00','13:00','ACTIVE']
    ]);
    sheet.setFrozenRows(1);
  }
  return book.getUrl();
}

function getShiftAppData() {
  const book = tkSpreadsheet_(), periods = tkPeriods_(book);
  if (!periods.length) throw new Error('No TK sheet with dated Time In and Time Out columns was found.');
  const today = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd');
  const current = periods.find(p => p.dateColumns.some(d => d.date === today)) || periods[periods.length - 1];
  return { today, todayAvailable: current.dateColumns.some(d => d.date === today),
    shifts: shiftOptions_(book), selectedSheet: current.sheetName,
    employees: employeeRows_(book.getSheetByName(current.sheetName)) };
}

function getEmployeesForPeriod(sheetName) {
  const book = tkSpreadsheet_();
  if (!tkPeriods_(book).some(p => p.sheetName === sheetName)) throw new Error('TK period not found.');
  return employeeRows_(book.getSheetByName(sheetName));
}

function processShift(input) {
  const employeeKey = String(input && input.employeeKey || '');
  const shiftId = String(input && input.shiftId || '');
  if (!employeeKey || !shiftId) throw new Error('Select an employee and shift.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const book = tkSpreadsheet_();
    const today = Utilities.formatDate(new Date(), 'Asia/Manila', 'yyyy-MM-dd');
    const matches = tkPeriods_(book).filter(p => p.dateColumns.some(d => d.date === today));
    if (!matches.length) throw new Error('No TK date column exists for today (' + today + '). Add today’s Time In/Out columns to the TK first.');
    if (matches.length > 1) throw new Error('More than one TK sheet contains today (' + today + '). Keep only one active sheet for this date.');
    const period = matches[0];
    const shift = shiftOptions_(book).find(s => s.id === shiftId);
    if (!shift) throw new Error('This shift is no longer ACTIVE. Refresh the app.');
    const sheet = book.getSheetByName(period.sheetName);
    const employee = employeeRows_(sheet).find(e => e.key === employeeKey);
    if (!employee) throw new Error('Employee is not on today’s TK sheet. Refresh and search again.');
    const dateColumn = period.dateColumns.find(d => d.date === today);
    const cells = sheet.getRange(employee.row, dateColumn.column, 1, 2);
    const values = cells.getValues()[0], notes = cells.getNotes()[0];
    for (let i = 0; i < 2; i++) {
      if (values[i] !== '' && !String(notes[i] || '').startsWith('BESI SCHEDULED SHIFT')) {
        throw new Error(employee.name + ' already has an unmarked Time In/Out value for today. Review the cells before processing.');
      }
    }
    cells.setNumberFormat('hh:mm');
    cells.setValues([[timeValue_(shift.startTime), timeValue_(shift.endTime)]]);
    cells.setNotes([[
      'BESI SCHEDULED SHIFT: ' + shift.name + ' · ' + today + ' · Time In',
      'BESI SCHEDULED SHIFT: ' + shift.name + ' · ' + today + ' · Time Out'
    ]]);
    cells.setBackground('#fff2cc');
    SpreadsheetApp.flush();
    return { employee: employee.name, shift: shift.name, date: today,
      startTime: shift.startTime, endTime: shift.endTime,
      url: book.getUrl() + '#gid=' + sheet.getSheetId() };
  } finally { lock.releaseLock(); }
}

function tkSpreadsheet_() {
  const id = String(TK_SHIFT.spreadsheetId || '').trim();
  const book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Use a native Google Sheet and open this project from Extensions → Apps Script.');
  return book;
}

function tkPeriods_(book) {
  const periods = [];
  book.getSheets().forEach(sheet => {
    if (sheet.getName() === TK_SHIFT.optionsTab || sheet.getLastRow() < 6 || sheet.getLastColumn() < 4) return;
    const h = sheet.getRange(1,1,Math.min(15,sheet.getLastRow()),Math.min(80,sheet.getLastColumn())).getDisplayValues();
    const year = periodYear_(h), dates = [];
    for (let r = 0; r < h.length-1; r++) for (let c = 0; c < h[r].length-1; c++) {
      const date = dateHeading_(h[r][c], year);
      if (date && /^TIME\s*IN$/i.test(String(h[r+1][c]).trim()) &&
          /^TIME\s*OUT$/i.test(String(h[r+1][c+1]).trim())) dates.push({ date, column:c+1 });
    }
    if (dates.length && nameHeader_(h)) {
      dates.sort((a,b)=>a.date.localeCompare(b.date));
      periods.push({sheetName:sheet.getName(),firstDate:dates[0].date,
        lastDate:dates[dates.length-1].date,dateColumns:dates});
    }
  });
  return periods.sort((a,b)=>a.lastDate.localeCompare(b.lastDate));
}

function periodYear_(headers) {
  for (const row of headers.slice(0,5)) for (const cell of row) {
    const match = String(cell||'').match(/\b(20\d{2})\b/);
    if (match) return Number(match[1]);
  }
  return Number(Utilities.formatDate(new Date(),'Asia/Manila','yyyy'));
}

function dateHeading_(value,year) {
  const text = String(value||'').trim().toUpperCase().replace(/SEPT/g,'SEP');
  let m = text.match(/^(\d{1,2})[-\s](JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)(?:[-\s](\d{4}))?$/);
  if (m) {
    const months=['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    return isoDate_(m[3]?Number(m[3]):year,months.indexOf(m[2])+1,Number(m[1]));
  }
  m=text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  return m?isoDate_(Number(m[1]),Number(m[2]),Number(m[3])):null;
}

function isoDate_(year,month,day) {
  const d=new Date(Date.UTC(year,month-1,day));
  if(d.getUTCFullYear()!==year||d.getUTCMonth()!==month-1||d.getUTCDate()!==day)return null;
  return [year,String(month).padStart(2,'0'),String(day).padStart(2,'0')].join('-');
}

function nameHeader_(headers) {
  for(let r=0;r<headers.length;r++)for(let c=0;c<headers[r].length;c++){
    if(/^(EMPLOYEE NAME|COMPLETE NAME)$/i.test(String(headers[r][c]).trim()))return {row:r+1,col:c};
  }
  return null;
}

function employeeRows_(sheet) {
  const h=sheet.getRange(1,1,Math.min(15,sheet.getLastRow()),Math.min(50,sheet.getLastColumn())).getDisplayValues();
  const header=nameHeader_(h);
  if(!header||header.row>=sheet.getLastRow())return [];
  const first=header.row+1;
  const values=sheet.getRange(first,1,sheet.getLastRow()-first+1,header.col+1).getDisplayValues();
  return values.map((row,i)=>{
    const name=String(row[header.col]||'').trim().toUpperCase().replace(/\s+/g,' ');
    const id=header.col?String(row[header.col-1]||'').trim():'';
    return {key:id?'ID:'+id:'NAME:'+name,id,name,row:first+i};
  }).filter(p=>p.name&&!/^(TOTAL|SUBTOTAL)$/.test(p.name)).sort((a,b)=>a.name.localeCompare(b.name));
}

function shiftOptions_(book) {
  const sheet=book.getSheetByName(TK_SHIFT.optionsTab);
  if(!sheet)throw new Error('Run setupShiftSystem once to create SHIFT_OPTIONS.');
  const rows=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,5).getDisplayValues():[];
  const shifts=rows.filter(r=>String(r[4]).trim().toUpperCase()==='ACTIVE').map(r=>({
    id:String(r[0]).trim(),name:String(r[1]).trim(),
    startTime:normalizeTime_(r[2]),endTime:normalizeTime_(r[3])
  }));
  if(!shifts.length||shifts.some(s=>!s.id||!s.name||!s.startTime||!s.endTime||s.startTime===s.endTime))
    throw new Error('Add valid ACTIVE shifts to SHIFT_OPTIONS with ID, name, start and end time.');
  if(new Set(shifts.map(s=>s.id)).size!==shifts.length)throw new Error('Duplicate active Shift IDs in SHIFT_OPTIONS.');
  return shifts;
}

function normalizeTime_(value) {
  const t=String(value||'').trim().toUpperCase();
  let m=t.match(/^([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?$/);
  if(m)return String(Number(m[1])).padStart(2,'0')+':'+m[2];
  m=t.match(/^(1[0-2]|[1-9]):([0-5]\d)(?::[0-5]\d)?\s*([AP]M)$/);
  return m?String(Number(m[1])%12+(m[3]==='PM'?12:0)).padStart(2,'0')+':'+m[2]:null;
}

function timeValue_(hhmm) {
  const bits=hhmm.split(':').map(Number);
  return (bits[0]*60+bits[1])/1440;
}
