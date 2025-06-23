// ==============
// KONFIGURASI
// ==============
const SPREADSHEET_ID = "1LA1TlhgltT2bqSN3H-LYasq9PtInVlqq98VPru8txoo";
const DATA_ENTRY_SHEET_NAME = "Form2";
const APPROVED_DATA_SHEET_NAME = "Form3";
const CABANG_SHEET_NAME = "Cabang";
const LOGIN_LOG_SHEET_NAME = "Log Login";

// --- PUSAT KONFIGURASI NAMA KOLOM ---
// Pastikan tulisan di sini SAMA PERSIS dengan header di Google Sheet Anda
const COLUMN_NAMES = {
  STATUS: "Status",
  TIMESTAMP: "Timestamp",
  EMAIL_PEMBUAT: "Email_Pembuat",
  PEMBERI_PERSETUJUAN: "Pemberi Persetujuan",
  WAKTU_PERSETUJUAN: "Waktu Persetujuan",
  TANGGAL: "Tanggal"
};
// -----------------------------------------

const CREATOR_JABATAN = "BRANCH BUILDING SUPPORT";
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
        const { parameter } = e;

        if (parameter.action === 'checkUserStatus') {
            if (!parameter.email) {
                return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Parameter email tidak ditemukan." })).setMimeType(ContentService.MimeType.JSON);
            }
            return checkUserLastSubmission(parameter.email);
        }

        const { action, row, approver } = parameter;
        if (!action || !row || !approver) {
            return createResponsePage({ title: 'Parameter Tidak Lengkap', message: 'Parameter URL tidak lengkap.', themeColor: '#dc3545', icon: '⚠' });
        }

        const lock = LockService.getScriptLock();
        lock.waitLock(30000);

        try {
            const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
            const statusColIndex = headers.indexOf(COLUMN_NAMES.STATUS);
            
            if (statusColIndex === -1) {
                return createResponsePage({ title: 'Kesalahan Konfigurasi', message: `Kolom status "${COLUMN_NAMES.STATUS}" tidak ditemukan.`, themeColor: '#dc3545', icon: '⚙️' });
            }
            if (parseInt(row) > SHEET.getLastRow()) {
                return createResponsePage({ title: 'Tindakan Sudah Diproses', message: 'Permintaan ini sudah diproses atau tidak ditemukan.', themeColor: '#ffc107', icon: 'ⓘ' });
            }
            const currentStatus = SHEET.getRange(parseInt(row), statusColIndex + 1).getValue();
            if (currentStatus !== 'Menunggu Persetujuan') {
                return createResponsePage({ title: 'Tindakan Sudah Diproses', message: `Permintaan ini sudah <strong>${currentStatus}</strong> sebelumnya.`, themeColor: '#ffc107', icon: 'ⓘ' });
            }
            
            const dataValues = SHEET.getRange(parseInt(row), 1, 1, SHEET.getLastColumn()).getValues()[0];
            const dataFromRowObj = {};
            headers.forEach((header, i) => { dataFromRowObj[header] = dataValues[i]; });

            const approverColIndex = headers.indexOf(COLUMN_NAMES.PEMBERI_PERSETUJUAN);
            const approvalTimeColIndex = headers.indexOf(COLUMN_NAMES.WAKTU_PERSETUJUAN);

            if (action === 'approve') {
                const approvedHeaders = APPROVED_SHEET.getRange(1, 1, 1, APPROVED_SHEET.getLastColumn()).getValues()[0];
                if (approvedHeaders.indexOf(COLUMN_NAMES.PEMBERI_PERSETUJUAN) === -1) APPROVED_SHEET.getRange(1, approvedHeaders.length + 1).setValue(COLUMN_NAMES.PEMBERI_PERSETUJUAN);
                if (approvedHeaders.indexOf(COLUMN_NAMES.WAKTU_PERSETUJUAN) === -1) APPROVED_SHEET.getRange(1, approvedHeaders.length + 2).setValue(COLUMN_NAMES.WAKTU_PERSETUJUAN);
                
                const finalApprovedHeaders = APPROVED_SHEET.getRange(1, 1, 1, APPROVED_SHEET.getLastColumn()).getValues()[0];
                const rowForApprovedSheet = finalApprovedHeaders.map(header => {
                    if (header === COLUMN_NAMES.STATUS) return 'Disetujui';
                    if (header === COLUMN_NAMES.PEMBERI_PERSETUJUAN) return approver;
                    if (header === COLUMN_NAMES.WAKTU_PERSETUJUAN) return new Date();
                    return dataFromRowObj[header] || '';
                });
                APPROVED_SHEET.appendRow(rowForApprovedSheet);
                SHEET.getRange(parseInt(row), statusColIndex + 1).setValue('Disetujui');

                if(approverColIndex > -1) SHEET.getRange(parseInt(row), approverColIndex + 1).setValue(approver);
                if(approvalTimeColIndex > -1) SHEET.getRange(parseInt(row), approvalTimeColIndex + 1).setValue(new Date());
                
                try { sendApprovalNotificationEmail(dataFromRowObj, approver); } catch (e) { console.error("Gagal kirim notif:", e); }
                return createResponsePage({ title: 'Persetujuan Berhasil', message: 'Tindakan Anda telah berhasil diproses.', themeColor: '#28a745', icon: '✔' });

            } else if (action === 'reject') {
                SHEET.getRange(parseInt(row), statusColIndex + 1).setValue('Ditolak');
                if(approverColIndex > -1) SHEET.getRange(parseInt(row), approverColIndex + 1).setValue(approver);
                if(approvalTimeColIndex > -1) SHEET.getRange(parseInt(row), approvalTimeColIndex + 1).setValue(new Date());
                
                try { sendRejectionNotificationEmail(dataFromRowObj, approver); } catch (e) { console.error("Gagal kirim notif tolak:", e); }
                return createResponsePage({ title: 'Permintaan Ditolak', message: 'Status permintaan telah diperbarui.', themeColor: '#dc3545', icon: '✖' });
            } else {
                return createResponsePage({ title: 'Aksi Tidak Valid', message: `Aksi "${action}" tidak dikenali.`, themeColor: '#dc3545', icon: '⚠' });
            }
        } finally {
            lock.releaseLock();
        }
    } catch (error) {
        console.error("Error in doGet:", error.toString(), error.stack);
        return createResponsePage({ title: 'Terjadi Kesalahan Internal', message: `Maaf, terjadi kesalahan.<br><small>Detail: ${error.message}</small>`, themeColor: '#dc3545', icon: '⚠' });
    }
};

const doPost = (request = {}) => {
    try {
        const { postData: { contents } = {} } = request;
        if (!contents) { throw new Error("Request tidak memiliki konten."); }
        const data = JSON.parse(contents);

        if (data.requestType === 'loginAttempt') {
            return logLoginAttempt(data);
        } else {
            const newRowIndex = appendToGoogleSheet(data);
            sendAutoEmail(data, newRowIndex);
            return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data submitted!" })).setMimeType(ContentService.MimeType.JSON);
        }
    } catch (error) {
        console.error(`Error in doPost:`, error.toString(), error.stack);
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `Kesalahan pada server: ${error.message}` })).setMimeType(ContentService.MimeType.JSON);
    }
};

function logLoginAttempt(data) {
    try {
        let logSheet = SPREADSHEET.getSheetByName(LOGIN_LOG_SHEET_NAME);
        if (!logSheet) {
            logSheet = SPREADSHEET.insertSheet(LOGIN_LOG_SHEET_NAME);
            logSheet.appendRow(['Timestamp', 'Username (Email)', 'Password (Cabang)', 'Status Login']);
        }
        const newRow = [new Date(), data.username || '', data.cabang || '', data.status || 'Unknown'];
        logSheet.appendRow(newRow);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Login attempt logged." })).setMimeType(ContentService.MimeType.JSON);
    } catch (e) {
        console.error("Gagal mencatat log login:", e.toString());
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `Gagal mencatat log: ${e.message}` })).setMimeType(ContentService.MimeType.JSON);
    }
}

function checkUserLastSubmission(email) {
    const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
    const emailColIndex = headers.indexOf(COLUMN_NAMES.EMAIL_PEMBUAT);
    const statusColIndex = headers.indexOf(COLUMN_NAMES.STATUS);

    if (emailColIndex === -1 || statusColIndex === -1) {
        const missing = [];
        if (emailColIndex === -1) missing.push(`'${COLUMN_NAMES.EMAIL_PEMBUAT}'`);
        if (statusColIndex === -1) missing.push(`'${COLUMN_NAMES.STATUS}'`);
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `Kolom ${missing.join(' dan ')} tidak ditemukan di sheet Form2.` })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const allData = SHEET.getDataRange().getValues();

    for (let i = allData.length - 1; i >= 1; i--) {
        const row = allData[i];
        if (row[emailColIndex] && row[emailColIndex].toString().trim() === email) {
            const lastStatus = row[statusColIndex] ? row[statusColIndex].toString().trim() : '';
            const responseData = { status: lastStatus, data: null };

            if (lastStatus === 'Ditolak') {
                const rowData = {};
                headers.forEach((header, index) => {
                    const key = header.replace(/ /g, "_");
                    rowData[key] = row[index];
                });
                responseData.data = rowData;
            }
            
            return ContentService.createTextOutput(JSON.stringify(responseData)).setMimeType(ContentService.MimeType.JSON);
        }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "No Data" })).setMimeType(ContentService.MimeType.JSON);
}

// =========================================================================
// FUNGSI INTI
// =========================================================================
function appendToGoogleSheet(data) {
    data[COLUMN_NAMES.TIMESTAMP] = new Date();
    const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
    if (headers.indexOf(COLUMN_NAMES.STATUS) === -1) throw new Error(`Kolom status "${COLUMN_NAMES.STATUS}" tidak ditemukan.`);
    
    const rowData = headers.map(header => {
        if (header === COLUMN_NAMES.STATUS) return 'Menunggu Persetujuan';
        if (data[header] instanceof Date) return data[header];
        if (header === COLUMN_NAMES.TANGGAL && typeof data[header] === 'string' && data[header].match(/^\d{4}-\d{2}-\d{2}$/)) {
          return new Date(data[header]);
        }
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
        console.warn(`No valid email recipients for branch: ${formData.Cabang}.`); return;
    }
    const approverEmail = recipients.to[0] || (recipients.cc[0] || 'unknown@example.com');
    const webAppUrl = ScriptApp.getService().getUrl();
    const approvalUrl = `${webAppUrl}?action=approve&row=${rowIndex}&approver=${encodeURIComponent(approverEmail)}`;
    const rejectionUrl = `${webAppUrl}?action=reject&row=${rowIndex}&approver=${encodeURIComponent(approverEmail)}`;
    const subject = `[PERLU PERSETUJUAN] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    const htmlForPdf = populateHtmlTemplate(formData);
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf).getAs('application/pdf').setName(`RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) { console.error("Error creating PDF blob:", e.message, e.stack); return; }
    const emailBodyHtml = `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:14px;line-height:1.6;}.approval-section{text-align:center;margin:30px 0;padding-top:20px;border-top:1px solid #eee;}.approval-button{text-decoration:none;color:#fff!important;padding:12px 25px;border-radius:5px;font-size:16px;font-weight:bold;margin:0 10px;display:inline-block;}.approve-btn{background-color:#28a745;}.reject-btn{background-color:#dc3545;}</style></head><body><p>Yth. Bapak/Ibu,</p><p>Dokumen <strong>Rekapitulasi Rencana Anggaran Biaya (RAB)</strong> untuk proyek <strong>${formData.Proyek || 'N/A'}</strong> memerlukan tinjauan dan persetujuan Anda.</p><p>Silakan periksa detailnya pada file PDF yang terlampir.</p><div class="approval-section"><p style="font-size:16px;font-weight:bold;">TINDAKAN PERSETUJUAN</p><p>Pilih salah satu tindakan di bawah ini:</p><br><a href="${approvalUrl}" class="approval-button approve-btn">SETUJUI</a><a href="${rejectionUrl}" class="approval-button reject-btn">TOLAK</a></div><p>Terima kasih.</p><p><em>--- Email ini dibuat secara otomatis.---</em></p></body></html>`;
    const options = { htmlBody: emailBodyHtml, attachments: [pdfBlob] };
    if (recipients.cc.length > 0) options.cc = recipients.cc.join(',');
    try {
        GmailApp.sendEmail(recipients.to.join(','), subject, "", options);
        console.log(`Email sent for project: ${formData.Proyek}`);
    } catch (e) { console.error(`Failed to send email: ${e.message}`, e.stack); }
}

function sendApprovalNotificationEmail(formData, approver) {
    const creatorEmails = getCreatorEmails(formData.Cabang);
    if (!creatorEmails || creatorEmails.length === 0) {
        console.warn(`Cannot send approval email: creator not found for branch ${formData.Cabang}.`); return;
    }
    const approverName = getNamaLengkapByEmail(approver);
    const approvalDetails = { email: approver, name: approverName };
    const subject = `[DISETUJUI] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    const htmlForPdf = populateHtmlTemplate(formData, approvalDetails);
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf).getAs('application/pdf').setName(`APPROVED_RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) { console.error("Failed to create PDF for approval notification:", e.message, e.stack); return; }
    const emailBodyHtml = `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:14px;}</style></head><body><p>Yth. Tim Branch Building Support,</p><p>Pengajuan RAB untuk proyek <strong>${formData.Proyek || 'N/A'}</strong> telah <strong>DISETUJUI</strong> oleh ${approver}.</p><p>Dokumen RAB final terlampir untuk arsip Anda.</p><br/><p>Terima kasih.</p><p><em>--- Email otomatis.---</em></p></body></html>`;
    const options = { htmlBody: emailBodyHtml, attachments: [pdfBlob] };
    try {
        const recipients = creatorEmails.join(',');
        GmailApp.sendEmail(recipients, subject, "", options);
        console.log(`Approval notification sent to (${recipients}) for project: ${formData.Proyek}`);
    } catch (e) { console.error(`Failed to send approval notification to (${creatorEmails.join(',')}) : ${e.message}`, e.stack); }
}

function sendRejectionNotificationEmail(formData, approver) {
    const creatorEmails = getCreatorEmails(formData.Cabang);
    if (!creatorEmails || creatorEmails.length === 0) {
        console.warn(`Cannot send rejection email: creator not found for branch ${formData.Cabang}.`); return;
    }
    const subject = `[PERLU REVISI] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    const htmlForPdf = populateHtmlTemplate(formData);
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf).getAs('application/pdf').setName(`REJECTED_RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) { console.error("Failed to create PDF for rejection notification:", e.message, e.stack); return; }
    const emailBodyHtml = `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:14px;}</style></head><body><p>Yth. Tim Branch Building Support,</p><p>Pengajuan RAB untuk proyek <strong>${formData.Proyek || 'N/A'}</strong> telah <strong>DITOLAK</strong> oleh ${approver} dan memerlukan revisi.</p><p>Dokumen asli terlampir sebagai referensi.</p><br/><p>Terima kasih.</p><p><em>--- Email otomatis.---</em></p></body></html>`;
    const options = { htmlBody: emailBodyHtml, attachments: [pdfBlob] };
    try {
        const recipients = creatorEmails.join(',');
        GmailApp.sendEmail(recipients, subject, "", options);
        console.log(`Rejection notification sent to (${recipients}) for project: ${formData.Proyek}`);
    } catch (e) { console.error(`Failed to send rejection notification to (${creatorEmails.join(',')}) : ${e.message}`, e.stack); }
}

// =========================================================================
// FUNGSI UTILITAS & TEMPLATE
// =========================================================================

function populateHtmlTemplate(data, approvalDetails = null) {
    const sipilCategories = ["PEKERJAAN PERSIAPAN", "PEKERJAAN BOBOKAN / BONGKARAN", "PEKERJAAN TANAH", "PEKERJAAN PONDASI & BETON", "PEKERJAAN PASANGAN", "PEKERJAAN BESI", "PEKERJAAN KERAMIK", "PEKERJAAN PLUMBING", "PEKERJAAN SANITARY & ACECORIES", "PEKERJAAN ATAP", "PEKERJAAN KUSEN, PINTU & KACA", "PEKERJAAN FINISHING", "PEKERJAAN TAMBAHAN"];
    const meCategories = ["INSTALASI", "FIXTURE"];
    const groupedItems = {};
    let grandTotalRp = 0;

    for (let i = 1; i <= 50; i++) {
        const jenisPekerjaan = data[`Jenis_Pekerjaan_${i}`];
        const kategoriPekerjaan = data[`Kategori_Pekerjaan_${i}`];
        if (jenisPekerjaan && kategoriPekerjaan) {
            if (!groupedItems[kategoriPekerjaan]) {
                groupedItems[kategoriPekerjaan] = [];
            }
            const item = {
                jenisPekerjaan: jenisPekerjaan,
                satuan: data[`Satuan_Item_${i}`],
                volume: data[`Volume_Item_${i}`],
                hargaMaterial: data[`Harga_Material_Item_${i}`],
                hargaUpah: data[`Harga_Upah_Item_${i}`],
                totalMaterial: data[`Total_Material_Item_${i}`],
                totalUpah: data[`Total_Upah_Item_${i}`],
                totalHarga: data[`Total_Harga_Item_${i}`],
            };
            groupedItems[kategoriPekerjaan].push(item);
            grandTotalRp += parseFloat(item.totalHarga) || 0;
        }
    }

    const lingkupPekerjaan = data.Lingkup_Pekerjaan;
    const categoriesToDisplay = (lingkupPekerjaan === 'Sipil') ? sipilCategories : (lingkupPekerjaan === 'ME') ? meCategories : Object.keys(groupedItems);
    
    let tablesHtml = '';
    categoriesToDisplay.forEach((category, categoryIndex) => {
        const items = groupedItems[category];
        if (items && items.length > 0) {
            let catSubTotalMaterial = 0, catSubTotalUpah = 0, catSubTotalHarga = 0;
            const itemRowsHtml = items.map((item, index) => {
                catSubTotalMaterial += parseFloat(item.totalMaterial) || 0;
                catSubTotalUpah += parseFloat(item.totalUpah) || 0;
                catSubTotalHarga += parseFloat(item.totalHarga) || 0;
                return `<tr><td>${index + 1}</td><td style="text-align: left;">${item.jenisPekerjaan || ''}</td><td>${item.satuan || ''}</td><td>${(parseFloat(item.volume) || 0).toFixed(2)}</td><td style="text-align: right;">${formatCurrency(item.hargaMaterial)}</td><td style="text-align: right;">${formatCurrency(item.hargaUpah)}</td><td style="text-align: right;">${formatCurrency(item.totalMaterial)}</td><td style="text-align: right;">${formatCurrency(item.totalUpah)}</td><td style="text-align: right;" class="total-amount-cell">${formatCurrency(item.totalHarga)}</td></tr>`;
            }).join('');
            const subTotalRowHtml = `<tr class="total-row"><td colspan="6" style="text-align: right; font-weight: bold;">SUB TOTAL</td><td style="text-align: right; font-weight: bold;">${formatCurrency(catSubTotalMaterial)}</td><td style="text-align: right; font-weight: bold;">${formatCurrency(catSubTotalUpah)}</td><td class="total-amount-cell" style="text-align: right; font-weight: bold;">${formatCurrency(catSubTotalHarga)}</td></tr>`;
            tablesHtml += `<h2 style="font-size: 14px; color: #333; border-bottom: 1px solid #ccc; padding-bottom: 5px; margin-top: 20px; margin-bottom: 10px; font-weight:bold;">${String.fromCharCode(65 + categoryIndex)}. ${category}</h2><div class="price-table-container"><table class="price-table"><thead><tr><th rowspan="2" style="width: 3%;">NO.</th><th rowspan="2">JENIS PEKERJAAN</th><th rowspan="2" style="width: 5%;">SATUAN</th><th style="width: 5%;">VOLUME</th><th colspan="2">HARGA SATUAN (Rp)</th><th colspan="2">TOTAL HARGA (Rp)</th><th style="width: 12%;">TOTAL HARGA (Rp)</th></tr><tr><th>a</th><th>Material (b)</th><th>Upah (c)</th><th>Material (d=a*b)</th><th>Upah (e=a*c)</th><th>(f=d+e)</th></tr></thead><tbody>${itemRowsHtml}${subTotalRowHtml}</tbody></table></div>`;
        }
    });

    const ppn = grandTotalRp * 0.11;
    const finalGrandTotal = grandTotalRp + ppn;
    const grandTotalHtml = `<table style="margin-top: 20px; width: 50%; float: right; border: none; font-size: 10px;"><tbody><tr class="total-row"><td style="border:none; text-align: right; font-weight: bold;">TOTAL (Rp)</td><td style="background-color:#e0ffff; border:1px solid #ddd; text-align: right; font-weight: bold;">${formatCurrency(grandTotalRp)}</td></tr><tr class="total-row"><td style="border:none; text-align: right; font-weight: bold;">PPN 11% (Rp)</td><td style="background-color:#e0ffff; border:1px solid #ddd; text-align: right; font-weight: bold;">${formatCurrency(ppn)}</td></tr><tr class="total-row"><td style="border:none; text-align: right; font-weight: bold;">GRAND TOTAL (Rp)</td><td style="background-color:#e0ffff; border:1px solid #ddd; text-align: right; font-weight: bold;">${formatCurrency(finalGrandTotal)}</td></tr></tbody></table><div style="clear: both;"></div>`;
    let template = `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;margin:0;padding:0;background-color:#f4f4f4;color:#333;width:210mm;min-height:297mm;margin:0 auto}.container{background-color:#fff;padding:10mm 8mm;border-radius:8px;width:100%;box-sizing:border-box;margin:0 auto;box-shadow:0 0 15px rgba(0,0,0,.1)}.header-company{text-align:left;margin-bottom:20px;border-bottom:1px solid #eee;padding-bottom:10px}.header-company p{margin:0;font-size:12px;color:#555;font-weight:700}.header-title{text-align:center;margin-bottom:30px}.header-title h1{font-size:19px;color:#333;margin:0}.section-info{margin-bottom:20px;font-size:11px}.section-info table{width:100%;border-collapse:collapse}.section-info table td{padding:5px 0;vertical-align:top}.section-info table td:first-child{width:180px;font-weight:700}.section-info table td:nth-child(2){width:10px;text-align:center}.section-info table td:last-child{text-align:left}.price-table-container{margin-top:0}.price-table{width:100%;border-collapse:collapse;font-size:10px}.price-table th,.price-table td{border:1px solid #ddd;padding:6px;text-align:center;white-space:nowrap}.price-table th{background-color:#e0ffff;font-weight:700;color:#555}.price-table td:nth-child(2){text-align:left;white-space:normal;width:auto}.price-table tbody tr td:last-child,.total-row .total-amount-cell{background-color:#e0ffff}.signatures{display:flex;justify-content:space-around;margin-top:60px;text-align:center; page-break-inside: avoid;}.signature-box{width:30%;display:flex;flex-direction:column;align-items:center}.signature-box p{margin:0;padding-top:5px;font-size:11px;color:#555}.signature-line{width:80%;height:1px;background-color:#333;margin-top:60px;margin-bottom:5px}</style></head><body><div class="container"><div class="header-company"><p>PT. SUMBER ALFARIA TRIJAYA, Tbk</p><p>BUILDING & MAINTENANCE DEPT</p><p>CABANG: <span id="branchName"></span></p></div><div class="header-title"><h1>REKAPITULASI RENCANA ANGGARAN BIAYA</h1></div><div class="section-info"><table><tr><td>LOKASI</td><td>:</td><td><span id="lokasi"></span></td></tr><tr><td>PROYEK</td><td>:</td><td><span id="proyek"></span></td></tr><tr><td>LINGKUP PEKERJAAN</td><td>:</td><td><span id="lingkupPekerjaan"></span></td></tr><tr><td>LUAS BANGUNAN</td><td>:</td><td><span id="luasBangunan"></span> m²</td></tr><tr><td>LUAS TERBANGUNAN</td><td>:</td><td><span id="luasTerbangunan"></span> m²</td></tr><tr><td>LUAS AREA TERBUKA/AREA PARKIR</td><td>:</td><td><span id="luasAreaTerbukaParkir"></span> m²</td></tr><tr><td>LUAS AREA SALES</td><td>:</td><td><span id="luasAreaSales"></span> m²</td></tr><tr><td>LUAS GUDANG</td><td>:</td><td><span id="luasGudang"></span> m²</td></tr><tr><td>TANGGAL RAB AWAL</td><td>:</td><td><span id="tanggalRabAwal"></span></td></tr><tr><td>WAKTU PELAKSANAAN</td><td>:</td><td><span id="waktuPelaksanaan"></span></td></tr></table></div>${tablesHtml}${grandTotalHtml}<div class="signatures"><div class="signature-box"><p>Dibuat</p><div class="signature-line"></div><p><span id="dibuatSignature"></span></p></div><div class="signature-box"><p>Mengetahui</p><div class="signature-line"></div><p><span id="mengetahuiSignature"></span></p></div><div class="signature-box"><p>Menyetujui</p><div class="signature-line"></div><p><span id="menyetujuiSignature"></span></p></div></div></div></body></html>`;
    
    const replaceContentById = (html, id, value) => html.replace(new RegExp(`(<span\\s+id="${id}">)[^<]*(<\\/span>)`), `$1${value || ''}$2`);
    template = replaceContentById(template, 'branchName', data.Cabang);
    template = replaceContentById(template, 'lokasi', data.Lokasi);
    template = replaceContentById(template, 'proyek', data.Proyek);
    template = replaceContentById(template, 'lingkupPekerjaan', data.Lingkup_Pekerjaan);
    template = replaceContentById(template, 'luasBangunan', data.Luas_Bangunan);
    template = replaceContentById(template, 'luasTerbangunan', data.Luas_Terbangunan);
    template = replaceContentById(template, 'luasAreaTerbukaParkir', data.Luas_Area_Terbuka_Area_Parkir);
    template = replaceContentById(template, 'luasAreaSales', data.Luas_Area_Sales);
    template = replaceContentById(template, 'luasGudang', data.Luas_Gudang);
    template = replaceContentById(template, 'tanggalRabAwal', data.Tanggal ? new Date(data.Tanggal).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '');
    template = replaceContentById(template, 'waktuPelaksanaan', data.Waktu_Pelaksanaan ? new Date(data.Waktu_Pelaksanaan).toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }) : '');
    template = replaceContentById(template, 'dibuatSignature', 'Br Building Support');
    template = replaceContentById(template, 'mengetahuiSignature', 'Br Building Coord');
    if (approvalDetails && approvalDetails.email) {
        const approverNameText = approvalDetails.name ? `<strong>${approvalDetails.name.toUpperCase()}</strong>` : '';
        const approvalSignature = `<span style="font-weight:normal; font-size: 9px; color: #28a745;">DISETUJUI oleh ${approvalDetails.email}</span><br/>${approverNameText}<br/>Br Build & Mtc Manager`;
        template = replaceContentById(template, 'menyetujuiSignature', approvalSignature);
    } else {
        template = replaceContentById(template, 'menyetujuiSignature', 'Br Build & Mtc Manager');
    }
    return template;
}

function getNamaLengkapByEmail(email) {
    if (!email) return "";
    const data = CABANG_SHEET.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('EMAIL_SAT'), namaCol = headers.indexOf('NAMA LENGKAP');
    if (emailCol === -1 || namaCol === -1) { console.error("Kolom 'EMAIL_SAT' atau 'NAMA LENGKAP' tidak ditemukan."); return ""; }
    for (let i = 1; i < data.length; i++) {
        const row = data[i], emailDiSheet = row[emailCol] ? row[emailCol].toString().trim().toLowerCase() : '';
        if (emailDiSheet === email.trim().toLowerCase()) return row[namaCol] ? row[namaCol].toString().trim() : '';
    }
    console.warn(`Nama lengkap untuk email ${email} tidak ditemukan.`); return "";
}

function getEmailRecipients(branchName) {
    const recipients = { to: [], cc: [] };
    if (!branchName) { console.error("getEmailRecipients dipanggil tanpa nama cabang."); return recipients; }
    const data = CABANG_SHEET.getDataRange().getValues(), headers = data[0];
    const branchNameCol = headers.indexOf('CABANG'), emailPicCol = headers.indexOf('EMAIL_SAT'), jabatanCol = headers.indexOf('JABATAN');
    if (branchNameCol === -1 || emailPicCol === -1 || jabatanCol === -1) { console.error(`Error: Kolom (CABANG, EMAIL_SAT, JABATAN) tidak ditemukan.`); return recipients; }
    for (let i = 1; i < data.length; i++) {
        const row = data[i], currentBranch = row[branchNameCol] ? row[branchNameCol].toString().trim().toLowerCase() : '', targetBranch = branchName.trim().toLowerCase();
        if (currentBranch === targetBranch) {
            const jabatan = row[jabatanCol] ? row[jabatanCol].toString().trim().toUpperCase() : '', email = row[emailPicCol] ? row[emailPicCol].toString().trim() : '';
            if (email) {
                if (jabatan === "BRANCH BUILDING & MAINTENANCE MANAGER") recipients.to.push(email);
                else if (jabatan === "BRANCH BUILDING COORDINATOR") recipients.cc.push(email);
            }
        }
    }
    if (recipients.to.length === 0) console.warn(`Penerima "TO" (Manager) tidak ditemukan untuk cabang: ${branchName}`);
    if (recipients.cc.length === 0) console.warn(`Penerima "CC" (Coordinator) tidak ditemukan untuk cabang: ${branchName}`);
    return recipients;
}

function getCreatorEmails(branchName) {
    const emails = [];
    if (!branchName) { console.error("getCreatorEmails dipanggil tanpa nama cabang."); return emails; }
    const data = CABANG_SHEET.getDataRange().getValues(), headers = data[0];
    const branchNameCol = headers.indexOf('CABANG'), emailPicCol = headers.indexOf('EMAIL_SAT'), jabatanCol = headers.indexOf('JABATAN');
    if (branchNameCol === -1 || emailPicCol === -1 || jabatanCol === -1) { console.error(`Error: Kolom 'CABANG', 'EMAIL_SAT', atau 'JABATAN' tidak ditemukan.`); return emails; }
    for (let i = 1; i < data.length; i++) {
        const row = data[i], currentBranch = row[branchNameCol] ? row[branchNameCol].toString().trim().toLowerCase() : '', targetBranch = branchName.trim().toLowerCase();
        if (currentBranch === targetBranch) {
            const jabatan = row[jabatanCol] ? row[jabatanCol].toString().trim().toUpperCase() : '';
            if (jabatan === CREATOR_JABATAN.toUpperCase()) {
                const email = row[emailPicCol] ? row[emailPicCol].toString().trim() : '';
                if (email) emails.push(email);
            }
        }
    }
    if (emails.length === 0) console.warn(`Tidak ada email yang ditemukan untuk jabatan "${CREATOR_JABATAN}" di cabang "${branchName}".`);
    return emails;
}

function formatCurrency(amount) {
    let num = typeof amount === 'number' ? amount : parseFloat(amount);
    if (isNaN(num)) return 'Rp 0';
    return 'Rp ' + num.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function createResponsePage(details) {
    const { title, message, themeColor, icon } = details;
    const logoUrl = 'https://upload.wikimedia.org/wikipedia/commons/9/9e/Alfamart_logo.svg';
    const html = `<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>${title}</title><style>body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif; background-color: #f0f2f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; color: #4b5563; } .card { background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); padding: 40px; text-align: center; max-width: 450px; width: 90%; border-top: 5px solid ${themeColor}; } .logo { max-width: 180px; margin-bottom: 24px; } .icon { font-size: 48px; line-height: 1; color: ${themeColor}; } h1 { font-size: 24px; font-weight: 600; margin-top: 16px; margin-bottom: 8px; color: #1f2937; } p { font-size: 16px; line-height: 1.6; margin-bottom: 24px; } .footer { font-size: 12px; color: #9ca3af; } small { color: #6b7280; } </style></head><body><div class="card"><img src="${logoUrl}" alt="Logo Alfamart" class="logo"><div class="icon">${icon}</div><h1>${title}</h1><p>${message}</p><div class="footer">Anda bisa menutup halaman ini.</div></div></body></html>`;
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}