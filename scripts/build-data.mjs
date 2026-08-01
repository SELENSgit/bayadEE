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

async function fetchExcelBuffer(startUrl) {
  const headers = {
    'User-Agent':
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
      '(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    Accept:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,*/*',
  };

  let currentUrl = startUrl;

  for (let hop = 0; hop < 5; hop++) {
    const res = await fetch(currentUrl, { redirect: 'follow', headers });

    if (!res.ok) {
      console.error(`Download failed: HTTP ${res.status} ${res.statusText}`);
      process.exit(1);
    }

    const buf = Buffer.from(await res.arrayBuffer());
    const magic = buf.subarray(0, 4).toString('hex');

    if (magic === '504b0304') {
      return buf; // real .xlsx (ZIP signature), we're done
    }

    // Not a real file — check if this is SharePoint's client-side
    // "Redirecting..." shell page and try to pull the real target URL out
    // of it (it embeds the destination in a script or meta refresh tag).
    const bodyText = buf.toString('utf8');
    const isRedirectShell = /Redirecting/i.test(bodyText) && /<html/i.test(bodyText);

    if (isRedirectShell) {
      const match =
        bodyText.match(/window\.location\.replace\(["']([^"']+)["']\)/i) ||
        bodyText.match(/window\.location\.href\s*=\s*["']([^"']+)["']/i) ||
        bodyText.match(/<meta[^>]+http-equiv=["']refresh["'][^>]+content=["'][^;]+;\s*url=([^"']+)["']/i) ||
        bodyText.match(/<a[^>]+id=["']download[^"']*["'][^>]+href=["']([^"']+)["']/i);

      if (match) {
        currentUrl = match[1].replace(/&amp;/g, '&');
        console.log(`Following redirect (hop ${hop + 1}) to: ${currentUrl.slice(0, 100)}...`);
        continue;
      }
    }

    // Couldn't find a next hop — dump debug info and give up.
    console.error(
      'The downloaded file is not a valid .xlsx, and no redirect target ' +
      'could be found in the response. This usually means the SharePoint ' +
      'link is not set to "Anyone with the link can view" — check the ' +
      'sharing settings on the Excel file.'
    );
    console.error(`Response content-type: ${res.headers.get('content-type')}`);
    console.error('First 1500 bytes of response body (for debugging):');
    console.error(bodyText.slice(0, 1500));
    process.exit(1);
  }

  console.error('Too many redirect hops without reaching a real .xlsx file.');
  process.exit(1);
}

const buf = await fetchExcelBuffer(url);

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
