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
            return createResponsePage({
                title: 'Parameter Tidak Lengkap',
                message: 'Aksi tidak dapat diproses karena parameter (action, row, approver) tidak lengkap pada URL.',
                themeColor: '#dc3545',
                icon: '⚠'
            });
        }

        const lock = LockService.getScriptLock();
        lock.waitLock(30000);

        try {
            const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
            const statusColIndex = headers.indexOf(STATUS_COLUMN_NAME);

            if (statusColIndex === -1) {
                return createResponsePage({
                    title: 'Kesalahan Konfigurasi',
                    message: `Kolom status dengan nama "${STATUS_COLUMN_NAME}" tidak dapat ditemukan di sheet "${DATA_ENTRY_SHEET_NAME}". Harap hubungi administrator.`,
                    themeColor: '#dc3545',
                    icon: '⚙️'
                });
            }
            
            if (parseInt(row) > SHEET.getLastRow()) {
                return createResponsePage({
                    title: 'Tindakan Sudah Diproses',
                    message: 'Permintaan ini sepertinya sudah diproses atau baris data tidak lagi ditemukan. Tidak ada tindakan lebih lanjut yang diperlukan.',
                    themeColor: '#ffc107',
                    icon: 'ⓘ'
                });
            }

            const currentStatus = SHEET.getRange(parseInt(row), statusColIndex + 1).getValue();
            if (currentStatus !== 'Menunggu Persetujuan') {
                return createResponsePage({
                    title: 'Tindakan Sudah Diproses',
                    message: `Permintaan ini sudah <strong>${currentStatus}</strong> sebelumnya. Tidak ada tindakan lebih lanjut yang diperlukan.`,
                    themeColor: '#ffc107',
                    icon: 'ⓘ'
                });
            }

            const dataFromRow = SHEET.getRange(parseInt(row), 1, 1, SHEET.getLastColumn()).getValues()[0];

            if (action === 'approve') {
                const approvedHeaders = APPROVED_SHEET.getRange(1, 1, 1, APPROVED_SHEET.getLastColumn()).getValues()[0];
                
                if (approvedHeaders.indexOf('Pemberi Persetujuan') === -1) APPROVED_SHEET.getRange(1, approvedHeaders.length + 1).setValue('Pemberi Persetujuan');
                if (approvedHeaders.indexOf('Waktu Persetujuan') === -1) APPROVED_SHEET.getRange(1, approvedHeaders.length + 2).setValue('Waktu Persetujuan');
                
                const finalApprovedHeaders = APPROVED_SHEET.getRange(1, 1, 1, APPROVED_SHEET.getLastColumn()).getValues()[0];

                const rowForApprovedSheet = finalApprovedHeaders.map(header => {
                    const index = headers.indexOf(header);
                    if (header === 'Status') return 'Disetujui';
                    if (header === 'Pemberi Persetujuan') return approver;
                    if (header === 'Waktu Persetujuan') return new Date();
                    return index > -1 ? dataFromRow[index] : '';
                });
                APPROVED_SHEET.appendRow(rowForApprovedSheet);

                SHEET.getRange(parseInt(row), statusColIndex + 1).setValue('Disetujui');
                SHEET.getRange(parseInt(row), statusColIndex + 2).setValue(approver); 
                SHEET.getRange(parseInt(row), statusColIndex + 3).setValue(new Date()); 
                
                try {
                    const approvedDataObj = {};
                    headers.forEach((header, index) => { approvedDataObj[header] = dataFromRow[index]; });
                    sendApprovalNotificationEmail(approvedDataObj, approver);
                } catch (emailError) {
                    console.error("Gagal mengirim email notifikasi persetujuan:", emailError.toString(), emailError.stack);
                }
                
                return createResponsePage({
                    title: 'Persetujuan Berhasil',
                    message: 'Tindakan Anda telah berhasil diproses. Notifikasi telah dikirimkan kepada pembuat permintaan.',
                    themeColor: '#28a745',
                    icon: '✔'
                });

            } else if (action === 'reject') {
                SHEET.getRange(parseInt(row), statusColIndex + 1).setValue('Ditolak');
                SHEET.getRange(parseInt(row), statusColIndex + 2).setValue(approver);
                SHEET.getRange(parseInt(row), statusColIndex + 3).setValue(new Date()); 
                
                try {
                    const rejectedDataObj = {};
                    headers.forEach((header, index) => { rejectedDataObj[header] = dataFromRow[index]; });
                    sendRejectionNotificationEmail(rejectedDataObj, approver);
                } catch (emailError) {
                    console.error("Gagal mengirim email notifikasi penolakan:", emailError.toString(), emailError.stack);
                }

                return createResponsePage({
                    title: 'Permintaan Ditolak',
                    message: 'Tindakan Anda telah berhasil diproses. Status permintaan telah diperbarui menjadi <strong>Ditolak</strong> dan notifikasi revisi telah dikirim.',
                    themeColor: '#dc3545',
                    icon: '✖'
                });

            } else {
                return createResponsePage({
                    title: 'Aksi Tidak Valid',
                    message: `Aksi "${action}" yang Anda berikan tidak dikenali oleh sistem.`,
                    themeColor: '#dc3545',
                    icon: '⚠'
                });
            }
        } finally {
            lock.releaseLock();
        }
    } catch (error) {
        console.error("Error in doGet:", error.toString(), error.stack);
        return createResponsePage({
            title: 'Terjadi Kesalahan Internal',
            message: `Maaf, terjadi kesalahan saat memproses permintaan Anda. Silakan hubungi administrator.<br><small>Detail: ${error.message}</small>`,
            themeColor: '#dc3545',
            icon: '⚠'
        });
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
    if (TIME_STAMP_COLUMN_NAME !== "") data[TIME_STAMP_COLUMN_NAME] = new Date();
    const headers = SHEET.getRange(1, 1, 1, SHEET.getLastColumn()).getValues()[0];
    if (headers.indexOf(STATUS_COLUMN_NAME) === -1) throw new Error(`Kolom status "${STATUS_COLUMN_NAME}" tidak ditemukan.`);
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
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf).getAs('application/pdf').setName(`RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) {
        console.error("Error creating PDF blob:", e.message, e.stack); return;
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

function populateHtmlTemplate(data, approvalDetails = null) {
    let template = `<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;margin:0;padding:0;background-color:#f4f4f4;color:#333;width:210mm;min-height:297mm;margin:0 auto}.container{background-color:#fff;padding:10mm 8mm;border-radius:8px;width:100%;box-sizing:border-box;margin:0 auto;box-shadow:0 0 15px rgba(0,0,0,.1)}.header-company{text-align:left;margin-bottom:20px;border-bottom:1px solid #eee;padding-bottom:10px}.header-company p{margin:0;font-size:9px;color:#555;font-weight:700}.header-title{text-align:center;margin-bottom:30px}.header-title h1{font-size:19px;color:#333;margin:0}.section-info{margin-bottom:20px;font-size:9px}.section-info table{width:100%;border-collapse:collapse}.section-info table td{padding:5px 0;vertical-align:top}.section-info table td:first-child{width:180px;font-weight:700}.section-info table td:nth-child(2){width:10px;text-align:center}.section-info table td:last-child{text-align:left}.price-table-container{margin-top:20px}.price-table{width:100%;border-collapse:collapse;font-size:8px}.price-table th,.price-table td{border:1px solid #ddd;padding:8px;text-align:center;white-space:nowrap}.price-table th{background-color:#e0ffff;font-weight:700;color:#555}.price-table td:nth-child(2){text-align:left;white-space:normal;width:80%}.price-table tbody tr td:first-child{width:5%}.price-table th:nth-child(3),.price-table td:nth-child(3),.price-table th:nth-child(4),.price-table td:nth-child(4),.price-table th:nth-child(5),.price-table td:nth-child(5){width:30%;text-align:right}.price-table tbody tr td:last-child,.total-row .total-amount-cell{background-color:#e0ffff}.notes{margin-top:30px;font-size:7px;color:#777;padding:10px;border-top:1px dashed #ddd}.signatures{display:flex;justify-content:space-around;margin-top:60px;text-align:center}.signature-box{width:30%;display:flex;flex-direction:column;align-items:center}.signature-box p{margin:0;padding-top:5px;font-size:8px;color:#555}.signature-line{width:80%;height:1px;background-color:#333;margin-top:60px;margin-bottom:5px}</style></head><body><div class="container"><div class="header-company"><p>PT. SUMBER ALFARIA TRIJAYA, Tbk</p><p>BUILDING & MAINTENANCE DEPT</p><p>CABANG: <span id="branchName"></span></p></div><div class="header-title"><h1>REKAPITULASI RENCANA ANGGARAN BIAYA</h1></div><div class="section-info"><table><tr><td>LOKASI</td><td>:</td><td><span id="lokasi"></span></td></tr><tr><td>PROYEK</td><td>:</td><td><span id="proyek"></span></td></tr><tr><td>LINGKUP PEKERJAAN</td><td>:</td><td><span id="lingkupPekerjaan"></span></td></tr><tr><td>LUAS BANGUNAN</td><td>:</td><td><span id="luasBangunan"></span> m²</td></tr><tr><td>LUAS TERBANGUNAN</td><td>:</td><td><span id="luasTerbangunan"></span> m²</td></tr><tr><td>LUAS AREA TERBUKA/AREA PARKIR</td><td>:</td><td><span id="luasAreaTerbukaParkir"></span> m²</td></tr><tr><td>LUAS AREA SALES</td><td>:</td><td><span id="luasAreaSales"></span> m²</td></tr><tr><td>LUAS GUDANG</td><td>:</td><td><span id="luasGudang"></span> m²</td></tr><tr><td>TANGGAL RAB AWAL</td><td>:</td><td><span id="tanggalRabAwal"></span></td></tr><tr><td>WAKTU PELAKSANAAN</td><td>:</td><td><span id="waktuPelaksanaan"></span></td></tr></table></div><div class="price-table-container"><table class="price-table"><thead><tr><th rowspan="2">NO.</th><th rowspan="2">JENIS PEKERJAAN</th><th colspan="3">Total Harga</th></tr><tr><th>Material<br>(a)</th><th>Upah<br>(b)</th><th>(Rp)<br>(c = a + b)</th></tr></thead><tbody>{{itemRows}}</tbody></table></div><p style="font-size:13px;margin-top:20px">Estimasi waktu pelaksanaan <span id="estimasiWaktuPelaksanaan"></span> hari, terhitung sejak SPK dikeluarkan</p><div class="signatures"><div class="signature-box"><p>Dibuat</p><div class="signature-line"></div><p><span id="dibuatSignature"></span></p></div><div class="signature-box"><p>Mengetahui</p><div class="signature-line"></div><p><span id="mengetahuiSignature"></span></p></div><div class="signature-box"><p>Menyetujui</p><div class="signature-line"></div><p><span id="menyetujuiSignature"></span></p></div></div><div class="notes"><p>Catatan:</p><p>Harga tersebut sesuai dengan Gambar Rencana Renovasi terlampir jika ada perubahan gambar dan spesifikasi material akan dilakukan perhitungan volume dari PT. Sumber Alfaria Trijaya, Tbk adalah sebagai referensi yang tidak mengikat dan kontraktor diwajibkan untuk mengecek ulang.</p><p>Reff: SAT/SOP/BDM/002 Prosedur Estimasi Biaya Renovasi</p></div></div></body></html>`;
    const replaceContentById = (html, id, value) => {
        return html.replace(new RegExp(`(<span\\s+id="${id}">)[^<]*(<\\/span>)`), `$1${value || ''}$2`);
    };
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
    
    let itemsData = [], subTotalMaterial = 0, subTotalUpah = 0, subTotalRp = 0;
    for (let i = 1; i <= 50; i++) {
        if (data[`Jenis_Pekerjaan_${i}`] && String(data[`Jenis_Pekerjaan_${i}`]).trim() !== '') {
            const material = parseFloat(data[`Total_Material_Item_${i}`] || 0), upah = parseFloat(data[`Total_Upah_Item_${i}`] || 0), totalHarga = parseFloat(data[`Total_Harga_Item_${i}`] || 0);
            itemsData.push({ jenisPekerjaan: data[`Jenis_Pekerjaan_${i}`], material, upah, totalHarga });
            subTotalMaterial += material; subTotalUpah += upah; subTotalRp += totalHarga;
        }
    }
    let itemRowsHtml = itemsData.map((item, index) => `<tr><td>${index+1}</td><td style="text-align: left;">${item.jenisPekerjaan||''}</td><td style="text-align: right;">${formatCurrency(item.material)}</td><td style="text-align: right;">${formatCurrency(item.upah)}</td><td style="text-align: right;" class="total-amount-cell">${formatCurrency(item.totalHarga)}</td></tr>`).join('');
    const ppn = subTotalRp * 0.11, grandTotal = subTotalRp + ppn;
    let totalsRowHtml = `<tr class="total-row"><td colspan="2" style="text-align: right;">SUB TOTAL (Rp)</td><td style="text-align: right;">${formatCurrency(subTotalMaterial)}</td><td style="text-align: right;">${formatCurrency(subTotalUpah)}</td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(subTotalRp)}</td></tr><tr class="total-row"><td colspan="2" style="text-align: right;">PEMBULATAN (Rp)</td><td colspan="2"></td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(0)}</td></tr><tr class="total-row"><td colspan="2" style="text-align: right;">PPN 11% (Rp)</td><td colspan="2"></td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(ppn)}</td></tr><tr class="total-row"><td colspan="2" style="text-align: right;">GRAND TOTAL (Rp)</td><td colspan="2"></td><td class="total-amount-cell" style="text-align: right;">${formatCurrency(grandTotal)}</td></tr>`;
    template = template.replace('{{itemRows}}', itemRowsHtml + totalsRowHtml);
    const tglAwal = new Date(data.Tanggal), tglAkhir = new Date(data.Waktu_Pelaksanaan);
    const estimasiHari = Math.ceil((tglAkhir - tglAwal) / (1000 * 60 * 60 * 24));
    template = replaceContentById(template, 'estimasiWaktuPelaksanaan', isNaN(estimasiHari) ? 'N/A' : estimasiHari);
    
    template = replaceContentById(template, 'dibuatSignature', 'Br Building Support');
    template = replaceContentById(template, 'mengetahuiSignature', 'Br Building Coord');
    if (approvalDetails && approvalDetails.email) {
        const approverNameText = approvalDetails.name ? `<strong>${approvalDetails.name.toUpperCase()}</strong>` : '';
        const approvalSignature = `<span style="font-weight:normal; font-size: 7px; color: #28a745;">DISETUJUI oleh ${approvalDetails.email}</span><br/>${approverNameText}<br/>Br Build & Mtc Manager`;
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
    if (emailCol === -1 || namaCol === -1) {
        console.error("Kolom 'EMAIL_SAT' atau 'NAMA LENGKAP' tidak ditemukan."); return "";
    }
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
    if (branchNameCol === -1 || emailPicCol === -1 || jabatanCol === -1) {
        console.error(`Error: Kolom (CABANG, EMAIL_SAT, JABATAN) tidak ditemukan.`); return recipients;
    }
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
    if (branchNameCol === -1 || emailPicCol === -1 || jabatanCol === -1) {
        console.error(`Error: Kolom 'CABANG', 'EMAIL_SAT', atau 'JABATAN' tidak ditemukan.`); return emails;
    }
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

function sendApprovalNotificationEmail(formData, approver) {
    const creatorEmails = getCreatorEmails(formData.Cabang);
    if (!creatorEmails || creatorEmails.length === 0) {
        console.warn(`Tidak dapat mengirim notifikasi persetujuan: email creator untuk cabang ${formData.Cabang} tidak ditemukan.`); return;
    }
    const approverName = getNamaLengkapByEmail(approver);
    const approvalDetails = { email: approver, name: approverName };
    const subject = `[DISETUJUI] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    const htmlForPdf = populateHtmlTemplate(formData, approvalDetails);
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf).getAs('application/pdf').setName(`APPROVED_RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) {
        console.error("Gagal membuat PDF untuk notifikasi:", e.message, e.stack); return;
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

/**
 * [FUNGSI BARU] Mengirim email notifikasi bahwa pengajuan DITOLAK.
 * @param {Object} formData - Data dari baris yang ditolak.
 * @param {string} approver - Email dari yang menolak.
 */
function sendRejectionNotificationEmail(formData, approver) {
    const creatorEmails = getCreatorEmails(formData.Cabang);
    if (!creatorEmails || creatorEmails.length === 0) {
        console.warn(`Tidak dapat mengirim notifikasi penolakan: email creator untuk cabang ${formData.Cabang} tidak ditemukan.`);
        return;
    }
    const subject = `[PERLU REVISI] RAB Proyek: ${formData.Proyek || 'N/A'}`;
    
    // Buat PDF asli tanpa tanda tangan digital.
    const htmlForPdf = populateHtmlTemplate(formData); // Dipanggil tanpa argumen kedua
    let pdfBlob;
    try {
        pdfBlob = HtmlService.createHtmlOutput(htmlForPdf)
            .getAs('application/pdf')
            .setName(`REJECTED_RAB_${formData.Proyek || 'NoProyek'}_${new Date().getTime()}.pdf`);
    } catch (e) {
        console.error("Gagal membuat PDF untuk notifikasi penolakan:", e.message, e.stack);
        return;
    }

    const emailBodyHtml = `
        <!DOCTYPE html>
        <html><head><style>body{font-family:Arial,sans-serif;font-size:14px;}</style></head>
        <body>
            <p>Yth. Tim Branch Building Support,</p>
            <p>
                Pengajuan Rencana Anggaran Biaya (RAB) untuk proyek <strong>${formData.Proyek || 'N/A'}</strong> 
                telah <strong>DITOLAK</strong> oleh ${approver} dan memerlukan revisi.
            </p>
            <p>Silakan periksa kembali data yang telah Anda kirimkan. Dokumen asli terlampir sebagai referensi.</p>
            <br/>
            <p>Terima kasih.</p>
            <p><em>--- Email ini dibuat secara otomatis oleh sistem. ---</em></p>
        </body></html>`;

    const options = { htmlBody: emailBodyHtml, attachments: [pdfBlob] };

    try {
        const recipients = creatorEmails.join(',');
        GmailApp.sendEmail(recipients, subject, "", options);
        console.log(`Email notifikasi penolakan berhasil dikirim ke (${recipients}) untuk proyek: ${formData.Proyek}`);
    } catch (e) {
        console.error(`Gagal mengirim email notifikasi penolakan ke (${creatorEmails.join(',')}) : ${e.message}`, e.stack);
    }
}


function formatCurrency(amount) {
    let num = typeof amount === 'number' ? amount : parseFloat(amount);
    if (isNaN(num)) return 'Rp 0';
    return 'Rp ' + num.toLocaleString('id-ID', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function createResponsePage(details) {
    const { title, message, themeColor, icon } = details;
    const logoUrl = 'https://commons.wikimedia.org/wiki/File:Alfamart_logo.svg';

    const html = `
    <!DOCTYPE html><html lang="id"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>${title}</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif; background-color: #f0f2f5; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; color: #4b5563; }
        .card { background-color: #ffffff; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); padding: 40px; text-align: center; max-width: 450px; width: 90%; border-top: 5px solid ${themeColor}; }
        .logo { max-width: 180px; margin-bottom: 24px; }
        .icon { font-size: 48px; line-height: 1; color: ${themeColor}; }
        h1 { font-size: 24px; font-weight: 600; margin-top: 16px; margin-bottom: 8px; color: #1f2937; }
        p { font-size: 16px; line-height: 1.6; margin-bottom: 24px; }
        .footer { font-size: 12px; color: #9ca3af; }
        small { color: #6b7280; }
    </style></head>
    <body><div class="card">
        <img src="${logoUrl}" alt="Logo Alfamart" class="logo">
        <div class="icon">${icon}</div><h1>${title}</h1><p>${message}</p>
        <div class="footer">Anda bisa menutup halaman ini.</div>
    </div></body></html>`;
    return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}