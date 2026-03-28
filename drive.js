'use strict';
const { google }   = require('googleapis');
const { Readable } = require('stream');
const fs   = require('fs');
const path = require('path');

const FOLDER_ID   = '1mmDo3SHRZGcClBSq5HNf_0MLM19-3Ywe';
const TOKEN_PATH  = path.join(__dirname, 'tokens.json');

let _drive = null;
const _ids = {}; // filename → fileId cache

// ── Auto-detect client_secret file ───────────────────────────────────────────
function _findClientSecret() {
  const files = fs.readdirSync(__dirname).filter(f =>
    f.startsWith('client_secret') && f.endsWith('.json')
  );
  return files.length ? path.join(__dirname, files[0]) : null;
}

// ── Init drive using OAuth2 tokens ───────────────────────────────────────────
function _getAuth() {
  if (_drive) return _drive;

  const secretPath = _findClientSecret();
  if (!secretPath || !fs.existsSync(TOKEN_PATH)) {
    throw new Error(
      '❌ ยังไม่ได้ตั้งค่า OAuth2\n' +
      '   รันคำสั่ง: node setup-oauth.js'
    );
  }

  const raw    = JSON.parse(fs.readFileSync(secretPath, 'utf8'));
  const creds  = raw.installed || raw.web;
  const tokens = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));

  const oAuth2 = new google.auth.OAuth2(
    creds.client_id, creds.client_secret, creds.redirect_uris[0]
  );
  oAuth2.setCredentials(tokens);

  // Auto-save refreshed tokens
  oAuth2.on('tokens', updated => {
    const current = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
    fs.writeFileSync(TOKEN_PATH, JSON.stringify({ ...current, ...updated }, null, 2));
  });

  _drive = google.drive({ version: 'v3', auth: oAuth2 });
  return _drive;
}

// ── Find file ID by name (cached) ─────────────────────────────────────────────
async function _findId(name) {
  if (_ids[name]) return _ids[name];
  const d = _getAuth();
  const r = await d.files.list({
    q: `name='${name}' and '${FOLDER_ID}' in parents and trashed=false`,
    fields: 'files(id)',
    spaces: 'drive',
    pageSize: 1,
  });
  if (r.data.files.length) {
    _ids[name] = r.data.files[0].id;
    return _ids[name];
  }
  return null;
}

// ── Read JSON from Drive ──────────────────────────────────────────────────────
async function readJSON(name) {
  const id = await _findId(name);
  if (!id) return null;
  const d = _getAuth();
  const r = await d.files.get({ fileId: id, alt: 'media' }, { responseType: 'json' });
  return r.data;
}

// ── Write JSON to Drive (create or update) ────────────────────────────────────
async function writeJSON(name, data) {
  const d       = _getAuth();
  const content = JSON.stringify(data);
  const id      = await _findId(name);

  if (id) {
    await d.files.update({
      fileId: id,
      media: { mimeType: 'application/json', body: Readable.from([content]) },
    });
  } else {
    const r = await d.files.create({
      requestBody: { name, parents: [FOLDER_ID], mimeType: 'application/json' },
      media:       { mimeType: 'application/json', body: Readable.from([content]) },
      fields: 'id',
    });
    _ids[name] = r.data.id;
  }
}

// ── Delete file from Drive ────────────────────────────────────────────────────
async function deleteFile(name) {
  const id = await _findId(name);
  if (!id) return;
  await _getAuth().files.delete({ fileId: id });
  delete _ids[name];
}

module.exports = { readJSON, writeJSON, deleteFile };
