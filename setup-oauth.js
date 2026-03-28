'use strict';
/**
 * รันครั้งเดียวเพื่อสร้าง tokens.json
 * node setup-oauth.js
 */
const { google }   = require('googleapis');
const readline     = require('readline');
const fs           = require('fs');
const path         = require('path');

const SCOPES = ['https://www.googleapis.com/auth/drive'];

function findClientSecret() {
  const files = fs.readdirSync(__dirname).filter(f =>
    f.startsWith('client_secret') && f.endsWith('.json')
  );
  if (!files.length) {
    console.error('❌ ไม่พบไฟล์ client_secret*.json ในโฟลเดอร์นี้');
    process.exit(1);
  }
  return path.join(__dirname, files[0]);
}

async function main() {
  const secretPath = findClientSecret();
  const raw        = JSON.parse(fs.readFileSync(secretPath, 'utf8'));
  const creds      = raw.installed || raw.web;

  const oAuth2 = new google.auth.OAuth2(
    creds.client_id,
    creds.client_secret,
    creds.redirect_uris[0]
  );

  const authUrl = oAuth2.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
    prompt: 'consent',
  });

  console.log('\n🔗 เปิดลิ้งค์นี้ในเบราว์เซอร์:');
  console.log('\n' + authUrl + '\n');

  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const code = await new Promise(resolve =>
    rl.question('📋 วาง code ที่ได้จากเบราว์เซอร์แล้วกด Enter: ', resolve)
  );
  rl.close();

  const { tokens } = await oAuth2.getToken(code.trim());
  fs.writeFileSync(path.join(__dirname, 'tokens.json'), JSON.stringify(tokens, null, 2));
  console.log('\n✅ บันทึก tokens.json เรียบร้อย! สามารถรัน server ได้เลย\n');
}

main().catch(e => { console.error('❌ Error:', e.message); process.exit(1); });
