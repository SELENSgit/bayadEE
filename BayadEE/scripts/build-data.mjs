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
    '"anyone with the link" sharing URL with &download=1 appended.'
  );
  process.exit(1);
}

const res = await fetch(url, { redirect: 'follow' });

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
  process.exit(1);
}

const workbook = XLSX.read(buf, { type: 'buffer', cellDates: true });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, {
  defval: '',
  raw: false,
  dateNF: 'yyyy-mm-dd',
});

function normKey(k) {
  return k.toString().trim().toLowerCase().replace(/[^a-z0-9]+/g, '_');
}

const records = rows
  .map((row) => {
    const norm = {};
    for (const [k, v] of Object.entries(row)) {
      norm[normKey(k)] = typeof v === 'string' ? v.trim() : v;
    }
    return {
      txnId: norm['txn_id'] || '',
      studentId: (norm['student_id'] || '').toString().trim(),
      fullName: norm['full_name'] || '',
      email: norm['email_address'] || '',
      event: norm['event_offense'] || norm['event'] || '',
      fineAmount: Number(norm['fine_amount']) || 0,
      date: norm['date'] || '',
      status: norm['status'] || '',
      datePaid: norm['date_paid'] || '',
    };
  })
  .filter((r) => r.studentId); // drop blank/empty rows

fs.mkdirSync('data', { recursive: true });
fs.writeFileSync(
  'data/fines.json',
  JSON.stringify(
    {
      updatedAt: new Date().toISOString(),
      records,
    },
    null,
    2
  )
);

console.log(`Wrote ${records.length} records to data/fines.json`);
