/**
 * ตัวอย่าง Google Apps Script สำหรับ Pink Bloom Hair Care
 * 1) สร้าง Google Sheet และแท็บชื่อ Orders
 * 2) Extensions > Apps Script แล้ววางโค้ดนี้
 * 3) Deploy > New deployment > Web app
 * 4) Execute as: Me; Who has access: เลือกตามนโยบายความปลอดภัยของคุณ
 *
 * คำเตือน: ตัวอย่างนี้ยังไม่ใช่ระบบ production และไม่มีระบบยืนยันตัวตนผู้ดูแล
 */
const SHEET_NAME = 'Orders';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('ไม่พบชีต Orders');
    if (!data.orderId || !data.customer || !data.phone || !data.address || !Array.isArray(data.items)) {
      throw new Error('ข้อมูลคำสั่งซื้อไม่ครบ');
    }
    data.items.forEach(item => {
      sheet.appendRow([
        data.orderId, data.createdAt || new Date(), data.customer, data.phone,
        data.address, data.note || '', item.sku, item.name, item.qty,
        item.price, item.subtotal, data.total, data.status || 'รอแพ็กสินค้า'
      ]);
    });
    return ContentService.createTextOutput(JSON.stringify({ok:true, orderId:data.orderId}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false, error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
