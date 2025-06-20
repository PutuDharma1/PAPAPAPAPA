// ==============
// KONFIGURASI
// ==============
const SPREADSHEET_ID = "1LA1TlhgltT2bqSN3H-LYasq9PtInVlqq98VPru8txoo"; // Ganti dengan ID Spreadsheet Anda
const DATA_ENTRY_SHEET_NAME = "Form2";
const APPROVED_DATA_SHEET_NAME = "Form3";
const CABANG_SHEET_NAME = "Cabang";
const TIME_STAMP_COLUMN_NAME = "Timestamp";
const STATUS_COLUMN_NAME = "Status";

// Konstanta untuk jabatan pembuat form
const CREATOR_JABATAN = "BRANCH BUILDING SUPPORT"; // Pastikan teks ini SAMA PERSIS dengan yang ada di sheet "Cabang"

const DEBUG_EMAIL_RECIPIENT = "";

// ==================
// VARIABEL GLOBAL & VALIDASI
// ==================
const SPREADSHEET = SpreadsheetApp.openById(SPREADSHEET_ID);
const SHEET = SPREADSHEET.getSheetByName(DATA_ENTRY_SHEET_NAME);
const APPROVED_SHEET = SPREADSHEET.getSheetByName(APPROVED_DATA_SHEET_NAME);
const CABANG_SHEET = SPREADSHEET.getSheetByName(CABANG_SHEET_NAME);

if (!SHEET) throw new Error(`Sheet dengan nama "${DATA_ENTRY_SHEET_NAME}" tidak ditemukan.`);
if (!APPROVED_SHEET) throw new Error(`Sheet dengan nama "${APPROVED_DATA_SHEET_NAME}" tidak ditemukan.`);
if (!CABANG_SHEET) throw new Error(`Sheet dengan nama "${CABANG_SHEET_NAME}" tidak ditemukan.`);


// =========================================================================
// FUNGSI UTAMA (doGet & doPost)
// =========================================================================

const doGet = (e) => {
    try {
        const { parameter: { action, row, approver } } = e;

        if (!action || !row || !approver) {
            return HtmlService.createHtmlOutput('<h1>Error</h1><p>Parameter tidak lengkap (action, row, approver). Aksi tidak dapat diproses.</p>');
        }

        const lock = LockService.getScriptLock();
        lock.waitLock(30000);

        try {
            const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
            const statusColIndex = headers.indexOf(STATUS_COLUMN_NAME);

            if (statusColIndex === -1) {
                return HtmlService.createHtmlOutput(`<h1>Error Konfigurasi</h1><p>Kolom "${STATUS_COLUMN_NAME}" tidak ditemukan di sheet "${DATA_ENTRY_SHEET_NAME}".</p>`);
            }
            
            if (parseInt(row) > SHEET.getLastRow()) {
                 return HtmlService.createHtmlOutput(`<h1>Tindakan Sudah Diproses</h1><p>Permintaan ini sepertinya sudah diproses atau baris data tidak lagi ditemukan.</p>`);
            }

            const currentStatus = SHEET.getRange(parseInt(row), statusColIndex + 1).getValue();
            if (currentStatus !== 'Menunggu Persetujuan') {
                return HtmlService.createHtmlOutput(`<h1>Tindakan Sudah Diproses</h1><p>Permintaan ini sudah <strong>${currentStatus}</strong> sebelumnya.</p>`);
            }

            if (action === 'approve') {
                const dataToCopy = SHEET.getRange(parseInt(row), 1, 1, SHEET.getLastColumn()).getValues()[0];
                const approvedHeaders = APPROVED_SHEET.getRange(1, 1, 1, APPROVED_SHEET.getLastColumn()).getValues()[0];
                
                if (approvedHeaders.indexOf('Pemberi Persetujuan') === -1) APPROVED_SHEET.getRange(1, approvedHeaders.length + 1).setValue('Pemberi Persetujuan');
                if (approvedHeaders.indexOf('Waktu Persetujuan') === -1) APPROVED_SHEET.getRange(1, approvedHeaders.length + 2).setValue('Waktu Persetujuan');
                
                const finalApprovedHeaders = APPROVED_SHEET.getRange(1, 1, 1, APPROVED_SHEET.getLastColumn()).getValues()[0];

                const rowForApprovedSheet = finalApprovedHeaders.map(header => {
                    const index = headers.indexOf(header);
                    if (header === 'Status') return 'Disetujui';
                    if (header === 'Pemberi Persetujuan') return approver;
                    if (header === 'Waktu Persetujuan') return new Date();
                    return index > -1 ? dataToCopy[index] : '';
                });
                APPROVED_SHEET.appendRow(rowForApprovedSheet);

                SHEET.getRange(parseInt(row), statusColIndex + 1).setValue('Disetujui');
                SHEET.getRange(parseInt(row), statusColIndex + 2).setValue(approver); 
                SHEET.getRange(parseInt(row), statusColIndex + 3).setValue(new Date()); 
                
                // --- Mengirim notifikasi email ke pembuat form ---
                try {
                    const approvedDataObj = {};
                    headers.forEach((header, index) => {
                        approvedDataObj[header] = dataToCopy[index];
                    });
                    sendApprovalNotificationEmail(approvedDataObj, approver);
                } catch (emailError) {
                    console.error("Gagal mengirim email notifikasi persetujuan:", emailError.toString(), emailError.stack);
                }
                // --- Akhir dari blok notifikasi ---

                return HtmlService.createHtmlOutput('<h1>Persetujuan Berhasil</h1><p>Tindakan telah berhasil diproses. Notifikasi telah dikirimkan kepada pembuat permintaan.</p>');

            } else if (action === 'reject') {
                SHEET.getRange(parseInt(row), statusColIndex + 1).setValue('Ditolak');
                SHEET.getRange(parseInt(row), statusColIndex + 2).setValue(approver);
                SHEET.getRange(parseInt(row), statusColIndex + 3).setValue(new Date()); 
                
                return HtmlService.createHtmlOutput('<h1>Permintaan Ditolak</h1><p>Tindakan telah berhasil diproses. Status telah diperbarui menjadi Ditolak.</p>');
            } else {
                 return HtmlService.createHtmlOutput(`<h1>Error</h1><p>Aksi "${action}" tidak valid.</p>`);
            }
        } finally {
            lock.releaseLock();
        }
    } catch (error) {
        console.error("Error in doGet:", error.toString(), error.stack);
        return HtmlService.createHtmlOutput(`<h1>Terjadi Error Internal</h1><p>Maaf, terjadi kesalahan saat memproses permintaan Anda. Silakan hubungi administrator.</p><p style="color:grey; font-size:10px;">Detail: ${error.message}</p>`);
    }
};


const doPost = (request = {}) => {
    try {
        const { postData: { contents } = {} } = request;
        if (!contents) {
            throw new Error("Request tidak memiliki konten (postData.contents is empty).");
        }
        const data = JSON.parse(contents);
        const newRowIndex = appendToGoogleSheet(data);
        sendAutoEmail(data, newRowIndex);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data submitted and email sent successfully!" })).setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        console.error(`Error in doPost for ${DATA_ENTRY_SHEET_NAME}:`, error.toString(), error.stack);
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `Terjadi kesalahan pada server: ${error.message}` })).setMimeType(ContentService.MimeType.JSON);
    }
};


// =========================================================================
// FUNGSI INTI (Append to Sheet, Send Email)
// =========================================================================

function appendToGoogleSheet(data) {
    if (TIME_STAMP_COLUMN_NAME !== "") {
        data[TIME_STAMP_COLUMN_NAME] = new Date();
    }
    const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
    if (headers.indexOf(STATUS_COLUMN_NAME) === -1) {
        throw new Error(`Kolom status "${STATUS_COLUMN_NAME}" tidak ditemukan.`);
    }
    const rowData = headers.map(header => {
        if (header === STATUS_COLUMN_NAME) return 'Menunggu Persetujuan';
        if (data[header] instanceof Date) return data[header];
        if (header === 'Tanggal' && typeof data[header] === 'string' && data[header].match(/^\d{4}-\d{2}-\d{2}$/)) return new Date(data[header]);
        return data[header] || "";
    });
    SHEET.appendRow(rowData);
    return SHEET.getLastRow();
}

function sendAutoEmail(formData, rowIndex) {
    let recipients = { to: [], cc: [] };
    if (DEBUG_EMAIL_RECIPIENT && DEBUG_EMAIL_RECIPIENT !== "") {
        recipients.to = [DEBUG_EMAIL_RECIPIENT];
    } else {
        recipients = getEmailRecipients(formData.Cabang);
    }
    if (recipients.to.length === 0 && recipients.cc.length === 0) {
        console.warn(`No valid email recipients found for branch: ${formData.Cabang}.`);
        return;
    }
    const approverEmail = recipients.to[0] || (recipients.cc[0] || 'unknown@example.com');
    const webAppUrl = ScriptApp.getService().getUrl();
    const approvalUrl = `${webAppUrl}?action=approve&row=${rowIndex}&approver=${encodeURIComponent(approverEmail)}`;
    const rejectionUrl = `${webAppUrl}?action=reject&row=${rowIndex}&approver=${encodeURIComponent(approverEmail)}`;
    const subject = `[PERLU PERSETUJUAN] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    const htmlForPdf = populateHtmlTemplate(formData);
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf)
            .getAs('application/pdf')
            .setName(`RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) {
        console.error("Error creating PDF blob:", e.message, e.stack);
        return;
    }
    const emailBodyHtml = `
        <!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:14px;line-height:1.6;}.approval-section{text-align:center;margin:30px 0;padding-top:20px;border-top:1px solid #eee;}.approval-button{text-decoration:none;color:#fff!important;padding:12px 25px;border-radius:5px;font-size:16px;font-weight:bold;margin:0 10px;display:inline-block;}.approve-btn{background-color:#28a745;}.reject-btn{background-color:#dc3545;}</style></head>
        <body><p>Yth. Bapak/Ibu,</p><p>Dokumen <strong>Rekapitulasi Rencana Anggaran Biaya (RAB)</strong> untuk proyek <strong>${formData.Proyek || 'N/A'}</strong> memerlukan tinjauan dan persetujuan Anda.</p><p>Silakan periksa detailnya pada file PDF yang terlampir dalam email ini.</p><div class="approval-section"><p style="font-size:16px;font-weight:bold;">TINDAKAN PERSETUJUAN</p><p>Setelah meninjau dokumen, silakan pilih salah satu tindakan di bawah ini:</p><br><a href="${approvalUrl}" class="approval-button approve-btn">SETUJUI</a><a href="${rejectionUrl}" class="approval-button reject-btn">TOLAK</a></div><p>Terima kasih.</p><p><em>--- Email ini dibuat secara otomatis oleh sistem. ---</em></p></body></html>`;
    const options = { htmlBody: emailBodyHtml, attachments: [pdfBlob] };
    if (recipients.cc.length > 0) options.cc = recipients.cc.join(',');
    try {
        GmailApp.sendEmail(recipients.to.join(','), subject, "", options);
        console.log(`Email sent successfully for project: ${formData.Proyek}`);
    } catch (e) {
        console.error(`Failed to send email: ${e.message}`, e.stack);
    }
}


// =========================================================================
// FUNGSI UTILITAS & TEMPLATE
// =========================================================================

function populateHtmlTemplate(data) {
    let template = `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;margin:0;padding:0;background-color:#f4f4f4;color:#333;width:210mm;min-height:297mm;margin:0 auto}.container{background-color:#fff;padding:10mm 8mm;border-radius:8px;width:100%;box-sizing:border-box;margin:0 auto;box-shadow:0 0 15px rgba(0,0,0,.1)}.header-company{text-align:left;margin-bottom:20px;border-bottom:1px solid #eee;padding-bottom:10px}.header-company p{margin:0;font-size:9px;color:#555;font-weight:700}.header-title{text-align:center;margin-bottom:30px}.header-title h1{font-size:19px;color:#333;margin:0}.section-info{margin-bottom:20px;font-size:9px}.section-info table{width:100%;border-collapse:collapse}.section-info table td{padding:5px 0;vertical-align:top}.section-info table td:first-child{width:180px;font-weight:700}.section-info table td:nth-child(2){width:10px;text-align:center}.section-info table td:last-child{text-align:left}.price-table-container{margin-top:20px}.price-table{width:100%;border-collapse:collapse;font-size:8px}.price-table th,.price-table td{border:1px solid #ddd;padding:8px;text-align:center;white-space:nowrap}.price-table th{background-color:#e0ffff;font-weight:700;color:#555}.price-table td:nth-child(2){text-align:left;white-space:normal;width:80%}.price-table tbody tr td:first-child{width:5%}.price-table th:nth-child(3),.price-table td:nth-child(3),.price-table th:nth-child(4),.price-table td:nth-child(4),.price-table th:nth-child(5),.price-table td:nth-child(5){width:30%;text-align:right}.price-table tbody tr td:last-child,.total-row .total-amount-cell{background-color:#e0ffff}.notes{margin-top:30px;font-size:7px;color:#777;padding:10px;border-top:1px dashed #ddd}.signatures{display:flex;justify-content:space-around;margin-top:60px;text-align:center}.signature-box{width:30%;display:flex;flex-direction:column;align-items:center}.signature-box p{margin:0;padding-top:5px;font-size:8px;color:#555}.signature-line{width:80%;height:1px;background-color:#333;margin-top:60px;margin-bottom:5px}</style></head><body><div class="container"><div class="header-company"><p>PT. SUMBER ALFARIA TRIJAYA, Tbk</p><p>BUILDING & MAINTENANCE DEPT</p><p>CABANG: <span id="branchName"></span></p></div><div class="header-title"><h1>REKAPITULASI RENCANA ANGGARAN BIAYA</h1></div><div class="section-info"><table><tr><td>LOKASI</td><td>:</td><td><span id="lokasi"></span></td></tr><tr><td>PROYEK</td><td>:</td><td><span id="proyek"></span></td></tr><tr><td>LINGKUP PEKERJAAN</td><td>:</td><td><span id="lingkupPekerjaan"></span></td></tr><tr><td>LUAS BANGUNAN</td><td>:</td><td><span id="luasBangunan"></span> m²</td></tr><tr><td>LUAS TERBANGUNAN</td><td>:</td><td><span id="luasTerbangunan"></span> m²</td></tr><tr><td>LUAS AREA TERBUKA/AREA PARKIR</td><td>:</td><td><span id="luasAreaTerbukaParkir"></span> m²</td></tr><tr><td>LUAS AREA SALES</td><td>:</td><td><span id="luasAreaSales"></span> m²</td></tr><tr><td>LUAS GUDANG</td><td>:</td><td><span id="luasGudang"></span> m²</td></tr><tr><td>TANGGAL RAB AWAL</td><td>:</td><td><span id="tanggalRabAwal"></span></td></tr><tr><td>WAKTU PELAKSANAAN</td><td>:</td><td><span id="waktuPelaksanaan"></span></td></tr></table></div><div class="price-table-container"><table class="price-table"><thead><tr><th rowspan="2">NO.</th><th rowspan="2">JENIS PEKERJAAN</th><th colspan="3">Total Harga</th></tr><tr><th>Material<br>(a)</th><th>Upah<br>(b)</th><th>(Rp)<br>(c = a + b)</th></tr></thead><tbody>{{itemRows}}</tbody></table></div><p style="font-size:13px;margin-top:20px">Estimasi waktu pelaksanaan <span id="estimasiWaktuPelaksanaan"></span> hari, terhitung sejak SPK dikeluarkan</p><div class="signatures"><div class="signature-box"><p>Dibuat</p><div class="signature-line"></div><p><span id="dibuatSignature">Br Building Support</span></p></div><div class="signature-box"><p>Mengetahui</p><div class="signature-line"></div><p><span id="mengetahuiSignature">Br Building Coord</span></p></div><div class="signature-box"><p>Menyetujui</p><div class="signature-line"></div><p><span id="menyetujuiSignature">Br Build & Mtc Manager</span></p></div></div><div class="notes"><p>Catatan:</p><p>Harga tersebut sesuai dengan Gambar Rencana Renovasi terlampir jika ada perubahan gambar dan spesifikasi material akan dilakukan perhitungan volume dari PT. Sumber Alfaria Trijaya, Tbk adalah sebagai referensi yang tidak mengikat dan kontraktor diwajibkan untuk mengecek ulang.</p><p>Reff: SAT/SOP/BDM/002 Prosedur Estimasi Biaya Renovasi</p></div></div></body></html>`;

    const replaceContentById = (html, id, value) => {
        const regex = new RegExp(`(<span\\s+id="${id}">)[^<]*(<\\/span>)`, 'g');
        return html.replace(regex, `$1${value !== undefined && value !== null ? value : ''}$2`);
    };
    template = replaceContentById(template, 'branchName', data.Cabang || '');
    template = replaceContentById(template, 'lokasi', data.Lokasi || '');
    template = replaceContentById(template, 'proyek', data.Proyek || '');
    template = replaceContentById(template, 'lingkupPekerjaan', data.Lingkup_Pekerjaan || '');
    template = replaceContentById(template, 'luasBangunan', data.Luas_Bangunan || '');
    template = replaceContentById(template, 'luasTerbangunan', data.Luas_Terbangunan || '');
    template = replaceContentById(template, 'luasAreaTerbukaParkir', data.Luas_Area_Terbuka_Area_Parkir || '');
    template = replaceContentById(template, 'luasAreaSales', data.Luas_Area_Sales || '');
    template = replaceContentById(template, 'luasGudang', data.Luas_Gudang || '');
    template = replaceContentById(template, 'tanggalRabAwal', data.Tanggal ? new Date(data.Tanggal).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '');
    template = replaceContentById(template, 'waktuPelaksanaan', data.Waktu_Pelaksanaan ? new Date(data.Waktu_Pelaksanaan).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '');
    
    let itemsData = [];
    let subTotalMaterial = 0, subTotalUpah = 0, subTotalRp = 0;
    for (let i = 1; i <= 50; i++) {
        if (data[`Jenis_Pekerjaan_${i}`] && String(data[`Jenis_Pekerjaan_${i}`]).trim() !== '') {
            const material = parseFloat(data[`Total_Material_Item_${i}`] || 0);
            const upah = parseFloat(data[`Total_Upah_Item_${i}`] || 0);
            const totalHarga = parseFloat(data[`Total_Harga_Item_${i}`] || 0);
            itemsData.push({ jenisPekerjaan: data[`Jenis_Pekerjaan_${i}`], material, upah, totalHarga });
            subTotalMaterial += material; subTotalUpah += upah; subTotalRp += totalHarga;
        }
    }
    let itemRowsHtml = itemsData.map((item, index) => `<tr><td>${index+1}</td><td style="text-align: left;">${item.jenisPekerjaan||''}</td><td style="text-align: right;">${formatCurrency(item.material)}</td><td style="text-align: right;">${formatCurrency(item.upah)}</td><td style="text-align: right;" class="total-amount-cell">${formatCurrency(item.totalHarga)}</td></tr>`).join('');
    const ppn = subTotalRp * 0.11;
    const grandTotal = subTotalRp + ppn;
    let totalsRowHtml = `<tr class="total-row"><td colspan="2" style="text-align: right;">SUB TOTAL (Rp)</td><td style="text-align: right;">${formatCurrency(subTotalMaterial)}</td><td style="text-align: right;">${formatCurrency(subTotalUpah)}</td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(subTotalRp)}</td></tr><tr class="total-row"><td colspan="2" style="text-align: right;">PEMBULATAN (Rp)</td><td colspan="2"></td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(0)}</td></tr><tr class="total-row"><td colspan="2" style="text-align: right;">PPN 11% (Rp)</td><td colspan="2"></td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(ppn)}</td></tr><tr class="total-row"><td colspan="2" style="text-align: right;">GRAND TOTAL (Rp)</td><td colspan="2"></td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(grandTotal)}</td></tr>`;
    template = template.replace('{{itemRows}}', itemRowsHtml + totalsRowHtml);

    const tglAwal = new Date(data.Tanggal);
    const tglAkhir = new Date(data.Waktu_Pelaksanaan);
    const estimasiHari = Math.ceil((tglAkhir - tglAwal) / (1000 * 60 * 60 * 24));
    template = replaceContentById(template, 'estimasiWaktuPelaksanaan', isNaN(estimasiHari) ? 'N/A' : estimasiHari);
    
    template = replaceContentById(template, 'dibuatSignature', 'Br Building Support');
    template = replaceContentById(template, 'mengetahuiSignature', 'Br Building Coord');
    template = replaceContentById(template, 'menyetujuiSignature', 'Br Build & Mtc Manager');

    return template;
}

/**
 * [FUNGSI YANG HILANG - DITAMBAHKAN KEMBALI]
 * Mengambil daftar email penerima (To/CC) untuk email PERMINTAAN PERSETUJUAN.
 * @param {string} branchName - Nama cabang untuk mencari penerima email.
 * @returns {{to: string[], cc: string[]}} Objek berisi array email untuk 'to' dan 'cc'.
 */
function getEmailRecipients(branchName) {
    const recipients = { to: [], cc: [] };
    if (!branchName) {
        console.error("getEmailRecipients dipanggil tanpa nama cabang.");
        return recipients;
    }

    const data = CABANG_SHEET.getDataRange().getValues();
    const headers = data[0];
    const branchNameCol = headers.indexOf('CABANG');
    const emailPicCol = headers.indexOf('EMAIL_SAT');
    const jabatanCol = headers.indexOf('JABATAN');

    if (branchNameCol === -1 || emailPicCol === -1 || jabatanCol === -1) {
        console.error(`Error: Kolom yang diperlukan (CABANG, EMAIL_SAT, JABATAN) tidak ditemukan di sheet "${CABANG_SHEET_NAME}".`);
        return recipients;
    }

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const currentBranch = row[branchNameCol] ? row[branchNameCol].toString().trim().toLowerCase() : '';
        const targetBranch = branchName.trim().toLowerCase();

        if (currentBranch === targetBranch) {
            const jabatan = row[jabatanCol] ? row[jabatanCol].toString().trim().toUpperCase() : '';
            const email = row[emailPicCol] ? row[emailPicCol].toString().trim() : '';
            if (email) {
                if (jabatan === "BRANCH BUILDING & MAINTENANCE MANAGER") {
                    recipients.to.push(email);
                } else if (jabatan === "BRANCH BUILDING COORDINATOR") {
                    recipients.cc.push(email);
                }
            }
        }
    }

    if (recipients.to.length === 0) {
       console.warn(`Penerima "TO" (Manager) tidak ditemukan untuk cabang: ${branchName}`);
    }
    if (recipients.cc.length === 0) {
       console.warn(`Penerima "CC" (Coordinator) tidak ditemukan untuk cabang: ${branchName}`);
    }
    return recipients;
}


/**
 * Mengambil SEMUA email pembuat form (Branch Building Support) berdasarkan nama cabang untuk NOTIFIKASI.
 * @param {string} branchName - Nama cabang.
 * @returns {string[]} Sebuah array/daftar email pembuat form. Akan kosong jika tidak ada yang ditemukan.
 */
function getCreatorEmails(branchName) {
    const emails = [];
    if (!branchName) {
        console.error("getCreatorEmails dipanggil tanpa nama cabang.");
        return emails;
    }

    const data = CABANG_SHEET.getDataRange().getValues();
    const headers = data[0];
    const branchNameCol = headers.indexOf('CABANG');
    const emailPicCol = headers.indexOf('EMAIL_SAT');
    const jabatanCol = headers.indexOf('JABATAN');

    if (branchNameCol === -1 || emailPicCol === -1 || jabatanCol === -1) {
        console.error(`Error: Kolom 'CABANG', 'EMAIL_SAT', atau 'JABATAN' tidak ditemukan di sheet "${CABANG_SHEET_NAME}".`);
        return emails;
    }

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const currentBranch = row[branchNameCol] ? row[branchNameCol].toString().trim().toLowerCase() : '';
        const targetBranch = branchName.trim().toLowerCase();

        if (currentBranch === targetBranch) {
            const jabatan = row[jabatanCol] ? row[jabatanCol].toString().trim().toUpperCase() : '';
            if (jabatan === CREATOR_JABATAN.toUpperCase()) {
                const email = row[emailPicCol] ? row[emailPicCol].toString().trim() : '';
                if (email) {
                    emails.push(email);
                }
            }
        }
    }

    if (emails.length === 0) {
        console.warn(`Tidak ada email yang ditemukan untuk jabatan "${CREATOR_JABATAN}" di cabang "${branchName}".`);
    }
    return emails;
}

/**
 * Mengirim email notifikasi bahwa pengajuan telah disetujui.
 * @param {Object} formData - Data dari baris yang disetujui.
 * @param {string} approver - Email dari pemberi persetujuan.
 */
function sendApprovalNotificationEmail(formData, approver) {
    const creatorEmails = getCreatorEmails(formData.Cabang);
    if (!creatorEmails || creatorEmails.length === 0) {
        console.warn(`Tidak dapat mengirim notifikasi persetujuan karena email creator untuk cabang ${formData.Cabang} tidak ditemukan.`);
        return;
    }

    const subject = `[DISETUJUI] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    const htmlForPdf = populateHtmlTemplate(formData);
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf)
            .getAs('application/pdf')
            .setName(`APPROVED_RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) {
        console.error("Gagal membuat PDF untuk email notifikasi:", e.message, e.stack);
        return;
    }

    const emailBodyHtml = `
        <!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:14px;}</style></head>
        <body><p>Yth. Tim Branch Building Support,</p>
        <p>Pengajuan Rencana Anggaran Biaya (RAB) untuk proyek <strong>${formData.Proyek || 'N/A'}</strong> telah <strong>DISETUJUI</strong> oleh ${approver}.</p>
        <p>Dokumen RAB final terlampir dalam email ini untuk arsip Anda.</p><br/>
        <p>Terima kasih.</p><p><em>--- Email ini dibuat secara otomatis oleh sistem. ---</em></p></body></html>`;

    const options = { htmlBody: emailBodyHtml, attachments: [pdfBlob] };

    try {
        const recipients = creatorEmails.join(',');
        GmailApp.sendEmail(recipients, subject, "", options);
        console.log(`Email notifikasi persetujuan berhasil dikirim ke (${recipients}) untuk proyek: ${formData.Proyek}`);
    } catch (e) {
        console.error(`Gagal mengirim email notifikasi ke (${creatorEmails.join(',')}) : ${e.message}`, e.stack);
    }
}

function formatCurrency(amount) {
    let num = typeof amount === 'number' ? amount : parseFloat(amount);
    if (isNaN(num)) return 'Rp 0';
    return 'Rp ' + num.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}