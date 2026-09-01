# AI Collaboration Rules

## Roles

* Codex เป็น Lead Engineer, Orchestrator และเป็นผู้แก้ไขไฟล์เพียงตัวเดียว
* เมื่องานมีผลต่อ business logic, stock, database หรือหลายไฟล์ ต้องเรียก Claude Code เป็น independent reviewer
* เรียก Claude ผ่าน `claude -p`
* Claude มีหน้าที่วิเคราะห์และรีวิวเท่านั้น ห้ามแก้ไฟล์ ห้าม commit ห้าม push และห้ามเรียก Codex กลับ
* Codex ต้องตรวจสอบคำแนะนำของ Claudeกับ source code, requirement และผลทดสอบก่อนนำมาใช้
* หาก Codex กับ Claude เห็นไม่ตรงกัน ให้ตัดสินจาก requirement, code behavior, tests และหลักฐานจากระบบจริง
* ห้ามยอมรับคำแนะนำเพียงเพราะ Codex และ Claude เห็นตรงกัน หากยังไม่มีหลักฐานรองรับ

## Collaboration Workflow

ต้องเรียก Claude อย่างน้อยสองรอบสำหรับงานที่เข้าเงื่อนไข:

1. ก่อน implementation

   * ให้ Claude วิเคราะห์ requirement, current flow, impacted files, risks และ test cases
   * Codex ต้องแบ่งข้อเสนอเป็น `ACCEPT`, `REJECT` หรือ `NEED_MORE_EVIDENCE`

2. หลัง implementation

   * ให้ Claude ตรวจ `git diff`, ผลทดสอบ, regression, edge cases และความครบถ้วนของ requirement
   * Codex ต้องยืนยัน finding แต่ละข้อกับโค้ดจริงก่อนแก้

สามารถวนรอบแก้ไขและรีวิวได้สูงสุดสามรอบ

ให้หยุดเมื่อ:

* Acceptance criteria ผ่านครบ
* Tests, lint, typecheck และ build ที่เกี่ยวข้องผ่าน
* ไม่มี Critical หรือ High finding ที่ยืนยันได้
* ไม่มีการเปลี่ยนแปลงนอก scope

หากครบสามรอบแล้วยังไม่ผ่าน ให้หยุดและรายงาน blocker พร้อมหลักฐาน ห้ามวนต่อโดยอัตโนมัติ

## Caveman

* ใช้ Caveman workflow กับงานวิเคราะห์ ออกแบบ แก้ไข และตรวจสอบโค้ด
* ก่อนเริ่มงาน ให้ตรวจว่า Caveman plugin หรือ skill พร้อมใช้งานและอ่านคำสั่งของ Caveman ที่เกี่ยวข้อง
* ปฏิบัติตาม workflow ของ Caveman ตลอดงาน ตราบใดที่ไม่ขัดกับ requirement หรือคำสั่งของผู้ใช้
* หาก Caveman ไม่พร้อมใช้งาน ให้แจ้งอย่างชัดเจนและดำเนินงานด้วย workflow ปกติ ห้ามอ้างว่าได้ใช้ Caveman แล้ว
* Caveman ไม่ได้แทนที่การตรวจสอบจาก Claude, tests หรือการตรวจ diff รอบสุดท้าย

## RTK

* ใช้ RTK สำหรับคำสั่ง shell ที่ RTK รองรับ เพื่อลด output และการใช้ token
* ก่อนเริ่มใช้ ให้ตรวจและอ่านคู่มือ RTK จาก `~/.codex/RTK.md` หากไฟล์มีอยู่
* ตัวอย่างคำสั่งที่ควรใช้ผ่าน RTK เมื่อรองรับ:

  * `rtk ls`
  * `rtk cat`
  * `rtk sed`
  * `rtk git status`
  * `rtk git diff`
* ห้ามบังคับใช้ RTK กับคำสั่งที่ไม่รองรับ คำสั่ง interactive หรือคำสั่งที่ RTK ทำให้ข้อมูลสำคัญถูกตัดออก
* หากต้องตรวจ output แบบสมบูรณ์เพื่อวิเคราะห์ error, test failure หรือ diff ให้ใช้คำสั่งปกติแทน RTK
* ให้เรียก `claude -p` โดยตรง เว้นแต่คู่มือ RTK ระบุอย่างชัดเจนว่ารองรับ Claude Code
* หาก RTK ไม่พร้อมใช้งาน ให้ใช้คำสั่งมาตรฐานต่อและรายงานสั้น ๆ ห้ามหยุดงานเพียงเพราะ RTK ใช้ไม่ได้

## Safety and Git

* Claude ห้ามแก้ไฟล์ ห้าม commit และห้าม push
* ห้ามให้ Claude เรียก Codex กลับ เพื่อป้องกัน recursive loop
* ห้ามให้ Codex และ Claude แก้ไฟล์เดียวกันพร้อมกัน
* ห้ามใช้ destructive Git operations
* ต้องรักษาการเปลี่ยนแปลงเดิมของผู้ใช้
* ห้าม commit หรือ push หากผู้ใช้ไม่ได้สั่ง
* ก่อนสรุปงาน Codex ต้องตรวจ `git status`, `git diff` และผลทดสอบด้วยตัวเอง
