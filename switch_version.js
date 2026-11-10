/**
 * Switch Version Utility: สลับเวอร์ชันระบบระหว่าง V5.3 (Claymorphism ดั้งเดิม) และ V5.4 (Winter Edition)
 * 
 * วิธีใช้งาน:
 *   node switch_version.js 5.3    -> สลับกลับไปเป็น Version 5.3 ทันที
 *   node switch_version.js 5.4    -> สลับมาเป็น Version 5.4 (ธีมหน้าหนาว) ทันที
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const targetVersion = process.argv[2];

if (!targetVersion || (targetVersion !== '5.3' && targetVersion !== '5.4')) {
  console.log('\n❌ กรุณาระบุเวอร์ชันที่ต้องการสลับ:');
  console.log('   node switch_version.js 5.3   (สลับกลับเป็น V5.3 Claymorphism เดิม)');
  console.log('   node switch_version.js 5.4   (สลับเป็น V5.4 Winter Edition หน้าหนาว)\n');
  process.exit(1);
}

const projectRoot = __dirname;
const backupFile = path.join(projectRoot, `index_v${targetVersion}_backup.html`);

if (!fs.existsSync(backupFile)) {
  console.error(`\n❌ ไม่พบไฟล์สำรองเวอร์ชัน ${targetVersion}: ${backupFile}\n`);
  process.exit(1);
}

const targetFiles = [
  path.join(projectRoot, 'index.html'),
  path.join(projectRoot, 'index_supabase.html'),
  path.join(projectRoot, 'GITHUB_DEPLOY', 'index.html')
];

const backupContent = fs.readFileSync(backupFile);
const hash = crypto.createHash('sha256').update(backupContent).digest('hex');

console.log(`\n======================================================`);
console.log(`🔄 กำลังสลับระบบไปยัง Version ${targetVersion}...`);
console.log(`======================================================`);

targetFiles.forEach(file => {
  fs.writeFileSync(file, backupContent);
  console.log(`✅ อัปเดตสำเร็จ: ${path.relative(projectRoot, file)}`);
});

console.log(`------------------------------------------------------`);
console.log(`✨ สลับระบบเป็น Version ${targetVersion} เรียบร้อยแล้ว 100%!`);
console.log(`🔑 SHA-256: ${hash}`);
console.log(`👉 ทดสอบผลลัพธ์ที่: http://localhost:3030/ (กด Ctrl + F5 หนึ่งครั้ง)`);
console.log(`======================================================\n`);
