// ==============
// CONFIGURATION
// ==============
const SPREADSHEET_ID = "116J074r3Nyy365G8_GXYoVzHGvqvjNvtESbPqQt5wUg"; // REMINDER: Update with your actual Spreadsheet ID
const DATA_ENTRY_SHEET_NAME = "Form2";
const CABANG_SHEET_NAME = "Cabang";
const TIME_STAMP_COLUMN_NAME = "Timestamp";

// Set this to an email address for debugging purposes.
// When set, all emails will be sent to this address.
// Set to null or an empty string "" to send to dynamic recipients from Cabang sheet.
const DEBUG_EMAIL_RECIPIENT = "iputudharma.puspa@gmail.com"; 

// ==================
// GLOBAL VARIABLES
// ==================
const SPREADSHEET = SpreadsheetApp.openById(SPREADSHEET_ID);
const SHEET = SPREADSHEET.getSheetByName(DATA_ENTRY_SHEET_NAME);
const CABANG_SHEET = SPREADSHEET.getSheetByName(CABANG_SHEET_NAME);

// HTML and CSS content for the email, with placeholders for data insertion
const EMAIL_HTML_TEMPLATE = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Rekapitulasi Rencana Anggaran Biaya</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f4f4f4;
            color: #333;
        }
        .container {
            background-color: #ffffff;
            padding: 30px;
            border-radius: 8px;
            max-width: 900px;
            margin: 0 auto;
            box-shadow: 0 0 15px rgba(0, 0, 0, 0.1);
        }
        .header-company {
            text-align: left;
            margin-bottom: 20px;
            border-bottom: 1px solid #eee;
            padding-bottom: 10px;
        }
        .header-company p {
            margin: 0;
            font-size: 14px;
            color: #555;
            font-weight: bold;
        }
        .header-title {
            text-align: center;
            margin-bottom: 30px;
        }
        .header-title h1 {
            font-size: 24px;
            color: #333;
            margin: 0;
        }
        .section-info {
            margin-bottom: 20px;
            font-size: 14px;
        }
        .section-info table {
            width: 100%;
            border-collapse: collapse;
        }
        .section-info table td {
            padding: 5px 0;
            vertical-align: top;
        }
        .section-info table td:first-child {
            width: 200px; /* Lebar kolom label disesuaikan */
            font-weight: bold;
        }
        .section-info table td:nth-child(2) {
            width: 10px;
            text-align: center;
        }
        .section-info table td:last-child {
            text-align: left;
        }
        .price-table-container {
            margin-top: 20px;
        }
        .price-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        .price-table th, .price-table td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: center;
        }
        .price-table th {
            background-color: #E0FFFF; /* Light Cyan */
            font-weight: bold;
            color: #555;
        }
        .price-table td:nth-child(2) { /* Jenis Pekerjaan */
            text-align: left;
        }
        .total-row td {
            font-weight: bold;
            background-color: #f9f9f9;
        }
        .total-row .label-normal-weight {
            font-weight: normal;
        }
        .total-row td:first-child {
            text-align: right;
            padding-right: 10px;
        }
        .price-table tbody tr td:last-child,
        .total-row .total-amount-cell {
            background-color: #E0FFFF; /* Light Cyan */
        }
        .notes {
            margin-top: 30px;
            font-size: 12px;
            color: #777;
            padding: 10px;
            border-top: 1px dashed #ddd;
        }
        .signatures {
            display: flex;
            justify-content: space-around;
            margin-top: 60px;
            text-align: center;
        }
        .signature-box {
            width: 30%;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .signature-box p {
            margin: 0;
            padding-top: 5px;
            font-size: 13px;
            color: #555;
        }
        .signature-line {
            width: 80%;
            height: 1px;
            background-color: #333;
            margin-top: 60px;
            margin-bottom: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-company">
            <p>PT. SUMBER ALFARIA TRIJAYA, Tbk</p>
            <p>BUILDING & MAINTENANCE DEPT</p>
            <p>CABANG: <span id="branchName"></span></p>
        </div>

        <div class="header-title">
            <h1>REKAPITULASI RENCANA ANGGARAN BIAYA</h1>
        </div>

        <div class="section-info">
            <table>
                <tr><td>LOKASI</td><td>:</td><td><span id="lokasi"></span></td></tr>
                <tr><td>PROYEK</td><td>:</td><td><span id="proyek"></span></td></tr>
                <tr><td>LINGKUP PEKERJAAN</td><td>:</td><td><span id="lingkupPekerjaan"></span></td></tr>
                <tr><td>LUAS BANGUNAN</td><td>:</td><td><span id="luasBangunan"></span> m²</td></tr>
                <tr><td>LUAS TERBANGUNAN</td><td>:</td><td><span id="luasTerbangunan"></span> m²</td></tr>
                <tr><td>LUAS AREA TERBUKA/AREA PARKIR</td><td>:</td><td><span id="luasAreaTerbukaParkir"></span> m²</td></tr>
                <tr><td>LUAS AREA SALES</td><td>:</td><td><span id="luasAreaSales"></span> m²</td></tr>
                <tr><td>LUAS GUDANG</td><td>:</td><td><span id="luasGudang"></span> m²</td></tr>
                <tr><td>TANGGAL RAB AWAL</td><td>:</td><td><span id="tanggalRabAwal"></span></td></tr>
                <tr><td>WAKTU PELAKSANAAN</td><td>:</td><td><span id="waktuPelaksanaan"></span></td></tr>
            </table>
        </div>

        <div class="price-table-container">
            <table class="price-table">
                <thead>
                    <tr>
                        <th rowspan="2">NO.</th>
                        <th rowspan="2">JENIS PEKERJAAN</th>
                        <th colspan="3">Total Harga</th>
                    </tr>
                    <tr>
                        <th>Material<br/>(a)</th>
                        <th>Upah<br/>(b)</th>
                        <th>(Rp)<br/>(c = a + b)</th>
                    </tr>
                </thead>
                <tbody>
                    <!--ITEM_ROWS_PLACEHOLDER-->
                    <tr class="total-row">
                        <td colspan="2" style="text-align: right;" class="label-normal-weight">SUB TOTAL (Rp)</td>
                        <td id="subTotalMaterial"></td>
                        <td id="subTotalUpah"></td>
                        <td class="total-amount-cell" id="subTotalRp"></td>
                    </tr>
                    <tr class="total-row">
                        <td colspan="2" style="text-align: right;" class="label-normal-weight">PEMBULATAN (Rp)</td>
                        <td colspan="2"></td>
                        <td class="total-amount-cell" id="pembulatan"></td>
                    </tr>
                    <tr class="total-row">
                        <td colspan="2" style="text-align: right;">PPN 11% (Rp)</td>
                        <td colspan="2"></td>
                        <td class="total-amount-cell" id="ppn"></td>
                    </tr>
                    <tr class="total-row">
                        <td colspan="2" style="text-align: right;" class="label-normal-weight">GRAND TOTAL (Rp)</td>
                        <td colspan="2"></td>
                        <td class="total-amount-cell" id="grandTotal"></td>
                    </tr>
                </tbody>
            </table>
        </div>

        <p style="font-size: 13px; margin-top: 20px;">
            Estimasi waktu pelaksanaan <span id="estimasiWaktuPelaksanaan"></span> hari, terhitung sejak SPK dikeluarkan
        </p>

        <div class="signatures">
            <div class="signature-box">
                <p>Dibuat</p>
                <div class="signature-line"></div>
                <p><span id="dibuatSignature">Br Building Support</span></p>
            </div>
            <div class="signature-box">
                <p>Mengetahui</p>
                <div class="signature-line"></div>
                <p><span id="mengetahuiSignature">Br Building Coord</span></p>
            </div>
            <div class="signature-box">
                <p>Menyetujui</p>
                <div class="signature-line"></div>
                <p><span id="menyetujuiSignature">Br Build & Mtc Manager</span></p>
            </div>
        </div>

        <div class="notes">
            <p>Catatan:</p>
            <p>Harga tersebut sesuai dengan Gambar Rencana Renovasi terlampir jika ada perubahan gambar dan spesifikasi material akan dilakukan perhitungan volume dari PT. Sumber Alfaria Trijaya, Tbk adalah sebagai referensi yang tidak mengikat dan kontraktor diwajibkan untuk mengecek ulang.</p>
            <p>Reff: SAT/SOP/BDM/002 Prosedur Estimasi Biaya Renovasi</p>
        </div>
    </div>
</body>
</html>`;

/**
 * Handles POST requests to the web app.
 * This function receives data submitted from a form or external source.
 * @param {object} request The request object.
 * @return {GoogleAppsScript.Content.TextOutput} A JSON response.
 */
const doPost = (request = {}) => {
  try {
    const { postData: { contents } = {} } = request;

    console.log("Received POST request with contents:", contents);
    const data = JSON.parse(contents);
    console.log("Parsed data:", data);

    appendToGoogleSheet(data);
    sendAutoEmail(data);

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data submitted and email sent successfully!" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error(`Error in doPost for ${DATA_ENTRY_SHEET_NAME}:`, error.toString(), error.stack);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message || "An unknown error occurred." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
};

/**
 * Appends data to the specified Google Sheet.
 * @param {object} data The data object to append.
 */
function appendToGoogleSheet(data) {
  if (TIME_STAMP_COLUMN_NAME !== "") {
    data[TIME_STAMP_COLUMN_NAME] = new Date();
  }

  if (!SHEET) {
    throw new Error(`Target sheet "${DATA_ENTRY_SHEET_NAME}" not found in the spreadsheet.`);
  }

  const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
  console.log("Sheet Headers:", headers);

  const rowData = headers.map(headerFld => {
    // Handle date formatting for 'Timestamp' and 'Tanggal'
    if (headerFld === TIME_STAMP_COLUMN_NAME && data[headerFld] instanceof Date) {
      return data[headerFld];
    }
    if (headerFld === 'Tanggal' && typeof data[headerFld] === 'string' && data[headerFld].match(/^\d{4}-\d{2}-\d{2}$/)) {
      return new Date(data[headerFld]);
    }
    return data[headerFld] || "";
  });
  console.log("Row Data to append:", rowData);

  SHEET.appendRow(rowData);
  console.log("Data appended successfully.");
}

/**
 * Sends an automated email with a PDF attachment generated from the HTML template.
 * @param {object} formData The data object submitted from the form.
 */
function sendAutoEmail(formData) {
  if (!CABANG_SHEET) {
    console.error(`Cabang sheet "${CABANG_SHEET_NAME}" not found.`);
    return;
  }

  let recipients = { to: [], cc: [] };

  if (DEBUG_EMAIL_RECIPIENT && DEBUG_EMAIL_RECIPIENT !== "") {
    recipients.to = [DEBUG_EMAIL_RECIPIENT];
    console.log("Sending email to debug recipient:", DEBUG_EMAIL_RECIPIENT);
  } else {
    recipients = getEmailRecipients(formData.Cabang);
    console.log("Dynamic email recipients:", recipients);
  }

  if (recipients.to.length === 0 && recipients.cc.length === 0) {
    console.warn(`No valid email recipients found for branch: ${formData.Cabang}. Email not sent.`);
    return;
  }

  const subject = `Rekapitulasi RAB Proyek: ${formData.Proyek || 'N/A'}`;
  let populatedHtml;
  try {
      populatedHtml = populateEmailTemplate(formData);
      console.log("HTML template populated successfully.");
  } catch (e) {
      console.error("Error populating email template:", e.message, e.stack);
      return; // Stop if template population fails
  }

  let pdfBlob;
  try {
      pdfBlob = HtmlService.createHtmlOutput(populatedHtml)
          .getAs('application/pdf') // Use getAs for more reliable PDF conversion
          .setName(`RAB_${formData.Proyek || 'NoProyek'}.pdf`);
      console.log("PDF blob created successfully.");
  } catch (e) {
      console.error("Error creating PDF blob:", e.message, e.stack);
      return; // Stop if PDF creation fails
  }

  const options = {
    htmlBody: "Salam,<br><br>Terlampir adalah dokumen Rekapitulasi Rencana Anggaran Biaya.<br><br>Terima kasih.",
    attachments: [pdfBlob],
  };

  if (recipients.cc.length > 0) {
      options.cc = recipients.cc.join(',');
  }

  try {
    GmailApp.sendEmail(recipients.to.join(','), subject, 'Please see the attached PDF document.', options);
    console.log(`Email sent successfully for project: ${formData.Proyek} to ${recipients.to.join(',')}`);
  } catch (e) {
    console.error(`Failed to send email for project ${formData.Proyek}: ${e.message}`, e.stack);
  }
}

/**
 * Retrieves email recipients (To, Cc) from the 'Cabang' sheet.
 * @param {string} branchName The name of the branch to look up.
 * @returns {{to: string[], cc: string[]}} An object with 'to' and 'cc' arrays of email addresses.
 */
function getEmailRecipients(branchName) {
  const recipients = { to: [], cc: [] };
  if (!CABANG_SHEET || !branchName) return recipients;

  const data = CABANG_SHEET.getDataRange().getValues();
  const headers = data[0];
  const branchNameCol = headers.indexOf('Nama Cabang');
  const emailPicCol = headers.indexOf('Email PIC');
  const emailCcCol = headers.indexOf('Email CC');

  if (branchNameCol === -1 || emailPicCol === -1) {
    console.warn("Cabang sheet missing 'Nama Cabang' or 'Email PIC' headers.");
    return recipients;
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[branchNameCol] && row[branchNameCol].toString().trim().toLowerCase() === branchName.trim().toLowerCase()) {
      if (row[emailPicCol]) recipients.to.push(row[emailPicCol].toString().trim());
      if (emailCcCol > -1 && row[emailCcCol]) {
        const ccs = row[emailCcCol].toString().split(',').map(email => email.trim()).filter(email => email);
        recipients.cc.push(...ccs);
      }
      break; 
    }
  }
  return recipients;
}

/**
 * Populates the HTML template with data.
 * @param {object} data The data object from the form submission.
 * @returns {string} The fully populated HTML string.
 */
function populateEmailTemplate(data) {
  let template = EMAIL_HTML_TEMPLATE;

  const replaceContentById = (html, id, value) => {
    const escapedId = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`(<span\\s+id="${escapedId}">)[^<]*(<\\/span>)`, 'g');
    return html.replace(regex, `$1${value !== undefined && value !== null ? value : ''}$2`);
  };

  // Populate info fields
  template = replaceContentById(template, 'branchName', data.Cabang || '');
  template = replaceContentById(template, 'lokasi', data.Lokasi || '');
  template = replaceContentById(template, 'proyek', data.Proyek || '');
  template = replaceContentById(template, 'lingkupPekerjaan', data.Lingkup_Pekerjaan || '');
  template = replaceContentById(template, 'luasBangunan', data.Luas_Bangunan || '');
  template = replaceContentById(template, 'luasTerbangunan', data.Luas_Terbangunan || '');
  template = replaceContentById(template, 'luasAreaTerbukaParkir', data.Luas_Area_Terbuka_Area_Parkir || '');
  template = replaceContentById(template, 'luasAreaSales', data.Luas_Area_Sales || '');
  template = replaceContentById(template, 'luasGudang', data.Luas_Gudang || '');
  template = replaceContentById(template, 'tanggalRabAwal', data.Tanggal ? new Date(data.Tanggal).toLocaleDateString('id-ID', {day: '2-digit', month: 'long', year: 'numeric'}) : '');
  template = replaceContentById(template, 'waktuPelaksanaan', data.Waktu_Pelaksanaan ? new Date(data.Waktu_Pelaksanaan).toLocaleDateString('id-ID', {day: '2-digit', month: 'long', year: 'numeric'}) : '');

  // Dynamically build items array and calculate totals
  let itemsData = [];
  let subTotalMaterial = 0;
  let subTotalUpah = 0;
  let subTotalRp = 0;
  const MAX_ITEMS = 50; 

  for (let i = 1; i <= MAX_ITEMS; i++) {
    const jenisPekerjaanKey = `Jenis_Pekerjaan_${i}`;
    if (data[jenisPekerjaanKey] && String(data[jenisPekerjaanKey]).trim() !== '') {
      const material = parseFloat(data[`Total_Material_Item_${i}`] || 0);
      const upah = parseFloat(data[`Total_Upah_Item_${i}`] || 0);
      const totalHarga = parseFloat(data[`Total_Harga_Item_${i}`] || 0);

      itemsData.push({
        jenisPekerjaan: data[jenisPekerjaanKey],
        material: material,
        upah: upah,
        totalHarga: totalHarga
      });

      subTotalMaterial += material;
      subTotalUpah += upah;
      subTotalRp += totalHarga;
    }
  }

  // Populate price table with item rows
  let itemRowsHtml = '';
  itemsData.forEach((item, index) => {
    itemRowsHtml += `
        <tr>
            <td>${index + 1}</td>
            <td style="text-align: left;">${item.jenisPekerjaan || ''}</td>
            <td style="text-align: right;">${formatCurrency(item.material)}</td>
            <td style="text-align: right;">${formatCurrency(item.upah)}</td>
            <td style="text-align: right;" class="total-amount-cell">${formatCurrency(item.totalHarga)}</td>
        </tr>
    `;
  });
  
  // *** CORRECTED LINE: Insert dynamic rows using the placeholder ***
  template = template.replace('<!--ITEM_ROWS_PLACEHOLDER-->', itemRowsHtml);

  // Calculate totals
  const pembulatan = 0; // Assuming pembulatan is 0
  const ppn = subTotalRp * 0.11;
  const grandTotal = subTotalRp + pembulatan + ppn;

  // Populate totals in template
  template = template.replace(/(<td\s+id="subTotalMaterial">)[^<]*(<\/td>)/, `$1${formatCurrency(subTotalMaterial)}$2`);
  template = template.replace(/(<td\s+id="subTotalUpah">)[^<]*(<\/td>)/, `$1${formatCurrency(subTotalUpah)}$2`);
  template = replaceContentById(template, 'subTotalRp', formatCurrency(subTotalRp));
  template = replaceContentById(template, 'pembulatan', formatCurrency(pembulatan));
  template = replaceContentById(template, 'ppn', formatCurrency(ppn));
  template = replaceContentById(template, 'grandTotal', formatCurrency(grandTotal));

  // Populate other dynamic fields
  const tglAwal = new Date(data.Tanggal);
  const tglAkhir = new Date(data.Waktu_Pelaksanaan);
  const estimasiHari = Math.ceil((tglAkhir - tglAwal) / (1000 * 60 * 60 * 24));
  template = replaceContentById(template, 'estimasiWaktuPelaksanaan', isNaN(estimasiHari) ? 'N/A' : estimasiHari);
  
  // Signatures (can be made dynamic if needed)
  template = replaceContentById(template, 'dibuatSignature', 'Br Building Support');
  template = replaceContentById(template, 'mengetahuiSignature', 'Br Building Coord');
  template = replaceContentById(template, 'menyetujuiSignature', 'Br Build & Mtc Manager');

  return template;
}

/**
 * Formats a number as Indonesian Rupiah currency.
 * @param {number} amount The amount to format.
 * @returns {string} The formatted currency string.
 */
function formatCurrency(amount) {
  if (typeof amount !== 'number') {
    amount = parseFloat(amount);
  }
  if (isNaN(amount)) {
    return 'Rp 0';
  }
  // Use toLocaleString for proper formatting with dots
  return 'Rp ' + amount.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
