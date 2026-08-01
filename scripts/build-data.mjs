// Downloads the shared Excel file from SharePoint/OneDrive and converts it
// into data/fines.json, which the static site reads at runtime.
//
// Expected columns in the sheet (first sheet, first row = headers):
//   TXN ID | Student ID | Full Name | Email Address | Event / Offense |
//   Fine Amount | Date | Status | Date Paid

import * as XLSX from 'xlsx';
import fs from 'fs';

const url = process.env.EXCEL_URL;

if (!url) {
  console.error(
    'Missing EXCEL_URL. Set it as a GitHub Actions secret (Settings > ' +
    'Secrets and variables > Actions) pointing at your SharePoint ' +
    '"anyone with the link" sharing URL with ?download=1 appended.'
  );
  process.exit(1);
}

const res = await fetch(url, {
  redirect: 'follow',
  headers: {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
      '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    Accept:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*',
  },
});

if (!res.ok) {
  console.error(`Download failed: HTTP ${res.status} ${res.statusText}`);
  process.exit(1);
}

const buf = Buffer.from(await res.arrayBuffer());

// A real .xlsx file is a ZIP archive and always starts with the bytes
// "PK\x03\x04". If the SharePoint link isn't actually public, this request
// silently returns an HTML login/redirect page instead — catch that early
// with a clear error rather than writing garbage JSON.
const magic = buf.subarray(0, 4).toString('hex');
if (magic !== '504b0304') {
  console.error(
    'The downloaded file is not a valid .xlsx (got HTML/text instead). ' +
    'This usually means the SharePoint link is not set to "Anyone with the ' +
    'link can view" — check the sharing settings on the Excel file.'
  );
  console.error(`Response content-type: ${res.headers.get('content-type')}`);
  console.error('First 300 bytes of response body (for debugging):');
  console.error(buf.subarray(0, 300).toString('utf8'));
  process.exit(1);
}

const workbook = XLSX.read(buf, { type: 'buffer',
