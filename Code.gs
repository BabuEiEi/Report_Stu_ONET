/**
 * ระบบรายงานจำนวนผู้เข้าสอบ O-NET (Official Version 2026)
 * รองรับ: บันทึกข้อมูล, สร้างหัวตารางอัตโนมัติ, และส่งข้อมูลดิบเพื่อทำ Dashboard กราฟ
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบรายงานจำนวนผู้เข้าสอบ O-NET')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ฟังก์ชันสำหรับรับข้อมูลจากฟอร์มและบันทึกลง Spreadsheet
 */
function processForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Data_Reports'); 
    
    // 1. ตรวจสอบว่ามี Sheet หรือยัง ถ้าไม่มีให้สร้างใหม่พร้อมหัวตารางสวยงาม
    if (!sheet) {
      sheet = ss.insertSheet('Data_Reports');
      const headers = [
        'วัน-เวลาที่ส่ง',      // Col 0
        'รอบที่สอบ',           // Col 1
        'สนามสอบ/โรงเรียน',    // Col 2
        'ชื่อผู้รายงาน',        // Col 3
        'จำนวนห้อง',           // Col 4
        'ผู้มีสิทธิ์ (สทศ.2)',   // Col 5
        'เข้าสอบ (สทศ.2)',     // Col 6
        'ขาดสอบ',              // Col 7
        'Walk-in (มีเลข)',     // Col 8
        'Walk-in (ไม่มีเลข)',  // Col 9
        'รวมเข้าสอบทั้งหมด'     // Col 10
      ];
      sheet.appendRow(headers);
      
      // จัดรูปแบบหัวตาราง (สีน้ำเงิน Navy Blue ตัวหนังสือขาว)
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground('#1e3a8a')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setHorizontalAlignment('center')
           .setVerticalAlignment('middle');
      
      // ล็อกแถวแรก
      sheet.setFrozenRows(1);
    }

    // 2. คำนวณผลรวมสุทธิ (Total) ฝั่ง Server เพื่อความถูกต้องของข้อมูล
    const grandTotal = Number(formData.attended) + Number(formData.walkinWithId) + Number(formData.walkinNoId);
    
    // 3. บันทึกข้อมูลลงแถวใหม่
    sheet.appendRow([
      new Date(),             // Timestamp
      formData.round,
      formData.school,
      formData.reporter,
      formData.rooms,
      formData.totalEligible,
      formData.attended,
      formData.absent,
      formData.walkinWithId,
      formData.walkinNoId,
      grandTotal              // ยอดรวมที่คำนวณแล้ว
    ]);
    
    return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (error) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

/**
 * ฟังก์ชันดึงข้อมูลสำหรับ Dashboard
 * ส่งคืนข้อมูลดิบ (Array 2D) ทั้งหมดเพื่อให้หน้าเว็บนำไป Filter และสร้างกราฟเอง
 */
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Data_Reports');
  
  // กรณีไม่มี Sheet หรือมีแค่หัวตาราง ให้ส่งค่าว่างกลับไป
  if (!sheet || sheet.getLastRow() < 2) {
    return []; 
  }
  
  // ดึงข้อมูลทั้งหมดตั้งแต่แถวที่ 2 ถึงแถวสุดท้าย (ไม่เอาหัวตาราง)
  // getDisplayValues() จะดึงข้อมูลมาเป็น String ตามที่ตาเห็นใน Sheet
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getDisplayValues();
  
  return data;
}
