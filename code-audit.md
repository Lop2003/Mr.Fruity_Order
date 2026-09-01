# Audit: Apps Script ระบบสั่งซื้อและตัดสต๊อก

ไฟล์ที่ตรวจ: `code.gs`

## Critical

### 1. เติมสต๊อกเขียนคนละโครงสร้างกับระบบตัดสต๊อก/ใบซื้ออ่าน

- ฟังก์ชัน/บรรทัด: `processStockUpdate()` บรรทัด 774-823, `deductStockForCutoff()` บรรทัด 1495-1579, `updatePurchaseSummarySheet()` บรรทัด 1585-1754
- ความรุนแรง: critical
- ปัญหา: `processStockUpdate()` เขียนไป `stockSS.getSheets()[0]` และใช้คอลัมน์ A-D แต่ฟังก์ชันหลักที่ใช้สต๊อกอ่านจากแท็บรายเดือน `ของในสต็อก JAN/FEB/...` และอ่านข้อมูลจาก section วันที่ในคอลัมน์ B-D
- ผลกระทบ: แอดมินพิมพ์ "เติมสต๊อก" แล้วข้อมูลอาจไม่ถูกใช้จริงตอนออกใบซื้อหรือตัดสต๊อก ทำให้ยอดคงเหลือ/ยอดต้องซื้อผิด

### 2. ตัดรอบไม่ atomic และอาจหักสต๊อกซ้ำ

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 117-127, `deductStockForCutoff()` บรรทัด 1495-1579, `createCutoffRoundDivider()` บรรทัด 1759-1788
- ความรุนแรง: critical
- ปัญหา: flow ตัดรอบหักสต๊อกก่อน แล้วค่อยสร้าง divider รอบ ถ้าหักสำเร็จแต่สร้าง divider ล้มเหลว หรือมี request ตัดรอบพร้อมกัน ทั้งสอง request จะเห็นรอบยังเปิดอยู่
- ผลกระทบ: external stock อาจถูกหักซ้ำ แต่ order sheet ยังมีสถานะรอบไม่ตรงกับ stock

### 3. ยกเลิกออเดอร์หลังตัดรอบแล้วไม่คืนสต๊อก

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 184-188, `cancelOrder()` บรรทัด 1331-1405, `deductStockForCutoff()` บรรทัด 1564-1570
- ความรุนแรง: critical
- ปัญหา: `cancelOrder()` ลบรายการจากชีตออเดอร์ได้ทุกแถวของวันนั้น รวมถึงรอบที่ปิดแล้ว แต่ไม่มี logic คืนสต๊อกที่เคยหักตอนตัดรอบ
- ผลกระทบ: order sheet กับ external stock ไม่ตรงกันทันทีหลังยกเลิกออเดอร์เก่าที่ถูกตัดรอบแล้ว

### 4. ใบซื้อถูกสร้างก่อน external stock section ของวันนั้น

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 257-259, `logOrderToExternalStock()` บรรทัด 833-903, `updatePurchaseSummarySheet()` บรรทัด 1601-1630
- ความรุนแรง: critical
- ปัญหา: ตอนรับออเดอร์เรียก `updatePurchaseSummarySheet()` ก่อน `logOrderToExternalStock()` ถ้า external stock ยังไม่มี section วันที่นั้น ใบซื้อจะคำนวณเหมือนไม่มี stock แล้วค่อยสร้างวันที่ใน external stock ทีหลังโดยไม่ rebuild อีกครั้ง
- ผลกระทบ: ใบซื้อรอบแรกของวันอาจแสดงว่าต้องซื้อทั้งหมด แม้ stock จริงมีอยู่

### 5. วันที่ไม่ normalize ทุก path

- ฟังก์ชัน/บรรทัด: `parseOrderMessage()` บรรทัด 383 และ 404, `doPost()` บรรทัด 139-144
- ความรุนแรง: critical
- ปัญหา: ถ้าวันที่อยู่บรรทัดแรกใช้ `normalizeDelivDate()` แต่ถ้าวันที่อยู่บรรทัดถัดไปใช้แค่ `replace(/-/g, "/")` ส่วนคำสั่ง `อัปเดตสต๊อก` ก็รับวันที่ดิบโดยไม่ normalize
- ผลกระทบ: วันเดียวกันอาจกลายเป็น `07/07/69` กับ `7/7/69` ทำให้สร้าง/อ่านคนละชีต เช่น `ออเดอร์-07-07-69` กับ `ออเดอร์-7-7-69`

## Race Condition / Consistency Risk

### 6. มีค่า lock timeout แต่ไม่ได้ใช้ `LockService`

- ฟังก์ชัน/บรรทัด: `CONFIG.LOCK_TIMEOUT_MS` บรรทัด 22
- ความรุนแรง: critical
- ปัญหา: มี config สำหรับ lock แต่ทั้งไฟล์ไม่มีการเรียก `LockService`
- ผลกระทบ: ทุก operation ที่อ่านแล้วเขียน เช่น รับออเดอร์, ตัดรอบ, ยกเลิก, rebuild ใบซื้อ เสี่ยงชนกันทั้งหมด

### 7. รับออเดอร์พร้อมกันเสี่ยงแทรกแถวผิดและ renumber ผิด

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 257, `updateDeliverySheet()` บรรทัด 1118-1308
- ความรุนแรง: should-fix
- ปัญหา: `updateDeliverySheet()` อ่าน `lastRow`/section เดิม แล้วค่อย insert/append และ renumber ถ้า webhook หลายตัวเข้าพร้อมกัน ข้อมูลที่อ่านไว้จะ stale
- ผลกระทบ: รายการอาจแทรกผิด section, ลำดับร้านผิด, หรือร้านเดียวกันถูกสร้างเป็นหลาย section

### 8. ใบซื้อถูก clear/rewrite จากหลาย path โดยไม่มี lock

- ฟังก์ชัน/บรรทัด: `onEdit()` บรรทัด 707, `doPost()` บรรทัด 127/188/258, `updatePurchaseSummarySheet()` บรรทัด 1637-1736
- ความรุนแรง: should-fix
- ปัญหา: `updatePurchaseSummarySheet()` เริ่มจาก `clearContents()` และ `clearFormats()` แล้วค่อยเขียนใหม่ แต่ถูกเรียกได้จาก webhook และ onEdit
- ผลกระทบ: ถ้าทำงานซ้อนกัน อาจเหลือใบซื้อครึ่งทาง หรือ process ที่เก่ากว่า overwrite ผลลัพธ์ล่าสุด

### 9. Pending cancel เก็บ state ไม่พอและไม่มี lock

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 224-226 และ 176-185
- ความรุนแรง: should-fix
- ปัญหา: preview เก็บแค่ `storeName`/`deliveryDate` ไม่เก็บ snapshot ของ row/item ที่ผู้ใช้เห็น
- ผลกระทบ: ถ้ามีออเดอร์เพิ่ม/ลบก่อนผู้ใช้พิมพ์ `ยืนยัน 1,3` เลขรายการที่ยืนยันอาจไม่ใช่รายการเดิม

### 10. `logOrderToExternalStock()` สร้างวันที่แบบ read-then-write โดยไม่มี lock

- ฟังก์ชัน/บรรทัด: `logOrderToExternalStock()` บรรทัด 855-896
- ความรุนแรง: should-fix
- ปัญหา: ฟังก์ชันอ่านว่า date section มีหรือยัง แล้วค่อยเขียนวันที่ ถ้าออเดอร์แรกของวันเข้าพร้อมกัน อาจเขียนหัววันที่ซ้ำ
- ผลกระทบ: external stock มี section วันที่ซ้ำ ทำให้ `deductStockForCutoff()` และ `updatePurchaseSummarySheet()` อ่าน section ผิดหรืออ่านแค่ section แรก

## Data / Logic Bugs

### 11. LINE webhook ประมวลผลแค่ event แรก

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 47-58
- ความรุนแรง: should-fix
- ปัญหา: LINE webhook ส่ง `events` เป็น array ได้ แต่โค้ดใช้เฉพาะ `eventData.events[0]`
- ผลกระทบ: ถ้า LINE ส่งหลาย event ใน request เดียว event หลังๆ จะถูก ignore

### 12. ไม่มี guard เมื่อไม่มีชีต `Logs`

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 62-65
- ความรุนแรง: should-fix
- ปัญหา: `logSheet.appendRow()` ถูกเรียกทันทีโดยไม่เช็คว่า `logSheet` มีจริงไหม
- ผลกระทบ: ถ้า rename/delete ชีต `Logs` บอทจะพังตั้งแต่ต้น request แม้ข้อความจะ parse ได้

### 13. `logOrderToExternalStock()` ไม่ log รายการสินค้าเลย

- ฟังก์ชัน/บรรทัด: `logOrderToExternalStock(items, deliveryDateStr)` บรรทัด 833-903
- ความรุนแรง: should-fix
- ปัญหา: parameter `items` ไม่ถูกใช้ ฟังก์ชันแค่สร้าง/หาแถววันที่ใน external stock
- ผลกระทบ: ชื่อฟังก์ชันทำให้เข้าใจผิด และถ้าคาดหวังว่า external stock เก็บรายการออเดอร์ด้วย ตอนนี้ไม่ได้ทำ

### 14. คำสั่ง `เติมสต๊อก` อาจ overwrite ไม่ใช่เพิ่ม

- ฟังก์ชัน/บรรทัด: `processStockUpdate()` บรรทัด 816-823
- ความรุนแรง: should-fix
- ปัญหา: ถ้าพบสินค้าเดิม จะ `setValues([[finalName, amount, finalUnit, today]])` ทับจำนวนเดิม ไม่ได้นำไปบวกกับ stock เดิม
- ผลกระทบ: ถ้าผู้ใช้คาดว่า "เติม 5 กก" คือเพิ่มจากของเดิม ยอด stock จะผิดทันที

### 15. `getOrderRounds()` อาจสร้างรอบว่าง/เลขรอบคลาดเคลื่อน

- ฟังก์ชัน/บรรทัด: `getOrderRounds()` บรรทัด 1426-1459, `createCutoffRoundDivider()` บรรทัด 1765-1777
- ความรุนแรง: should-fix
- ปัญหา: `createCutoffRoundDivider()` เริ่ม `cutoffCount = 1` แล้ว `nextRound = cutoffCount + 1` ทำให้ divider แรกเป็น `รอบ 2` ส่วน `getOrderRounds()` push รอบปิดเมื่อเจอ divider และ push รอบล่าสุดว่างท้ายเสมอ
- ผลกระทบ: ข้อความ reply ใช้ `res.round - 1` เพื่อชดเชย ทำให้ logic อ่านยากและมีโอกาสคลาดเคลื่อนตอนมี divider ว่าง/ตัดรอบซ้ำ

### 16. `onEdit()` rebuild ใบซื้อจาก cache ทำให้ stock ล่าสุดไม่ถูกดึง

- ฟังก์ชัน/บรรทัด: `onEdit()` บรรทัด 699-707, `updatePurchaseSummarySheetFromCache()` บรรทัด 1581-1583
- ความรุนแรง: should-fix
- ปัญหา: เมื่อแก้ชีตออเดอร์ด้วยมือ จะ rebuild ใบซื้อแบบ `useCache=true` ซึ่งไม่อ่าน external stock ใหม่
- ผลกระทบ: ถ้ามีการแก้ออเดอร์หลัง stock เปลี่ยน ใบซื้ออาจใช้ stock เก่าจากประวัติ หรือแสดง stock เป็น 0 สำหรับรายการใหม่ที่ไม่มี history

### 17. Cache mapping อาจ stale ถ้า `onEdit` trigger ไม่ได้ติดตั้งหรือ cache ใหญ่เกิน

- ฟังก์ชัน/บรรทัด: `getMappingDictionary()` บรรทัด 725-768, `onEdit()` บรรทัด 689-692, `setupOnce()` บรรทัด 1074-1078
- ความรุนแรง: minor
- ปัญหา: cache ถูกล้างเมื่อ edit ชีต `Mapping` ผ่าน trigger เท่านั้น ถ้า trigger ไม่ถูก setup หรือแก้ผ่านบาง automation cache อาจอยู่จน TTL หมด
- ผลกระทบ: สินค้า/alias ที่แก้ใหม่อาจยังไม่ถูกใช้ทันที

## Dead Code / Legacy

### 18. `orderId` สร้างแล้วไม่ใช้

- ฟังก์ชัน/บรรทัด: `doPost()` บรรทัด 247-251
- ความรุนแรง: minor
- ปัญหา: สร้าง `orderId` แต่ไม่เขียนลงชีตหรือส่งต่อ
- ผลกระทบ: dead code และไม่มี id สำหรับ trace order จริง

### 19. Dashboard อ่านชีต legacy `Orders`/`Stock`

- ฟังก์ชัน/บรรทัด: `refreshDashboard()` บรรทัด 909-1047
- ความรุนแรง: should-fix
- ปัญหา: flow หลักเขียนไป `ออเดอร์-DD-MM-YY` และ external stock file แต่ dashboard อ่าน `Orders` และ `Stock`
- ผลกระทบ: dashboard อาจว่างหรือแสดงข้อมูลเก่าที่ไม่ตรงกับระบบจริง

### 20. `onEdit()` branch สำหรับ `Orders` น่าจะ legacy

- ฟังก์ชัน/บรรทัด: `onEdit()` บรรทัด 680-687
- ความรุนแรง: minor
- ปัญหา: flow หลักไม่ได้เขียนชีต `Orders` แล้ว
- ผลกระทบ: logic นี้อาจไม่ได้ใช้จริง และทำให้เข้าใจผิดว่ายังมีชีต order กลาง

### 21. `setupOnce()` สร้าง header ให้ชีต legacy

- ฟังก์ชัน/บรรทัด: `setupOnce()` บรรทัด 1051-1080
- ความรุนแรง: minor
- ปัญหา: setup ยังเตรียม `Orders` และ `Stock` แต่ flow ปัจจุบันใช้ชีตรายวันและ external stock รายเดือน
- ผลกระทบ: คนดูแลระบบอาจเข้าใจผิดว่าต้องใช้ชีต `Orders`/`Stock`

## Performance

### 22. `updateDeliverySheet()` เรียก Sheet API ในลูปต่อ item

- ฟังก์ชัน/บรรทัด: `updateDeliverySheet()` บรรทัด 1199-1221 และ 1264-1280
- ความรุนแรง: should-fix
- ปัญหา: path เพิ่มสินค้าใน section เดิมใช้ `getLastRow()`, `appendRow()`, `insertRowBefore()`, `getRange().set...` ซ้ำต่อรายการ และ path section ใหม่ยัง set format ต่อ row
- ผลกระทบ: ร้านที่สั่งหลายรายการจะช้า และเพิ่มโอกาส timeout ใน Apps Script

### 23. `deductStockForCutoff()` set stock ทีละแถว

- ฟังก์ชัน/บรรทัด: `deductStockForCutoff()` บรรทัด 1564-1570
- ความรุนแรง: should-fix
- ปัญหา: loop แล้ว `setValue()` ทีละสินค้า
- ผลกระทบ: ตัดรอบช้าเมื่อ stock section ใหญ่ ควร batch update คอลัมน์ C ของช่วงเดียว

### 24. `cancelOrder()` ลบแถวทีละแถว

- ฟังก์ชัน/บรรทัด: `cancelOrder()` บรรทัด 1385-1386
- ความรุนแรง: should-fix
- ปัญหา: `deleteRow()` ในลูปหลายครั้ง ถึงจะลบจากล่างขึ้นบนถูกแล้ว แต่ยังเป็น API call หลายครั้ง
- ผลกระทบ: ยกเลิกออเดอร์ใหญ่ช้า และเสี่ยง timeout

### 25. `logOrderToExternalStock()` อ่านทั้งคอลัมน์ `B:B`

- ฟังก์ชัน/บรรทัด: `logOrderToExternalStock()` บรรทัด 872
- ความรุนแรง: minor
- ปัญหา: `getRange("B:B").getValues()` อ่านทั้งคอลัมน์
- ผลกระทบ: external stock โตขึ้นแล้วจะช้าโดยไม่จำเป็น

### 26. `refreshDashboard()` set background ทีละแถว

- ฟังก์ชัน/บรรทัด: `refreshDashboard()` บรรทัด 1002-1005 และ 1034-1035
- ความรุนแรง: minor
- ปัญหา: หลัง batch `setValues()` แล้วยัง `setBackground()` ทีละ row
- ผลกระทบ: dashboard ช้าเมื่อจำนวนสินค้าหรือ mapping เยอะ

## ลำดับแก้ที่แนะนำ

1. ใส่ `LockService` ครอบ critical flow: รับออเดอร์, ตัดรอบ, ยกเลิก, เติมสต๊อก, rebuild ใบซื้อ
2. รวม source of truth ของ stock ให้เหลือที่เดียว: เลือก external monthly stock หรือชีต `Stock` แล้วปรับทุกฟังก์ชันให้อ่าน/เขียนที่เดียวกัน
3. ทำ date normalization ให้เข้าทุก path ก่อนสร้างชื่อชีตหรืออ่านชีต
4. แก้ flow รับออเดอร์ให้สร้าง/หา external stock section ก่อน rebuild ใบซื้อ
5. จำกัดการยกเลิกออเดอร์รอบที่ปิดแล้ว หรือเพิ่ม logic คืนสต๊อก
6. ลบ/แยก legacy `Orders`/`Stock`/Dashboard ถ้าไม่ได้ใช้แล้ว
