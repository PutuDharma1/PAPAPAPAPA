// Estimasi_rab/script.js

// --- Global Variable Declarations ---
let form;
let submitButton;
let messageDiv;
let grandTotalAmount;
let lingkupPekerjaanSelect;
let sipilTablesWrapper;
let meTablesWrapper;
let currentResetButton;
let categorizedPrices = {};
let pendingStoreCodes = [];
let approvedStoreCodes = [];

// --- Helper Functions ---
const formatRupiah = (number) => {
  return new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", minimumFractionDigits: 0, maximumFractionDigits: 0 }).format(number);
};

const parseRupiah = (formattedString) => {
  const cleanedString = String(formattedString).replace(/Rp\s?|\./g, "").replace(/,/g, ".");
  return parseFloat(cleanedString) || 0;
};

const populateJenisPekerjaanOptionsForNewRow = (rowElement) => {
  const category = rowElement.dataset.category;
  const scope = rowElement.dataset.scope;
  const searchInput = rowElement.querySelector(".jenis-pekerjaan-search-input");
  const hiddenSelect = rowElement.querySelector(".jenis-pekerjaan");
  const suggestionsList = rowElement.querySelector(".jenis-pekerjaan-suggestions");

  if (!searchInput || !hiddenSelect || !suggestionsList) {
    console.error("Error: Required elements not found in row:", rowElement);
    return;
  }

  let dataSource = {};
  if (scope === "Sipil" && categorizedPrices.categorizedSipilPrices) {
    dataSource = categorizedPrices.categorizedSipilPrices;
  } else if (scope === "ME" && categorizedPrices.categorizedMePrices) {
    dataSource = categorizedPrices.categorizedMePrices;
  } else {
    return;
  }

  const itemsInCategory = dataSource[category] || [];
  hiddenSelect.innerHTML = '<option value="">-- Pilih Jenis Pekerjaan --</option>';
  itemsInCategory.forEach((item) => {
    const option = document.createElement("option");
    option.value = item["Jenis Pekerjaan"];
    option.textContent = item["Jenis Pekerjaan"];
    hiddenSelect.appendChild(option);
  });

  const removeListeners = (inputEl) => {
    if (inputEl._inputHandler) inputEl.removeEventListener("input", inputEl._inputHandler);
    if (inputEl._focusHandler) inputEl.removeEventListener("focus", inputEl._focusHandler);
    if (inputEl._blurHandler) inputEl.removeEventListener("blur", inputEl._blurHandler);
  };
  removeListeners(searchInput);

  searchInput._inputHandler = () => {
    const searchTerm = searchInput.value.toLowerCase();
    suggestionsList.innerHTML = "";
    const itemsToDisplay = searchTerm.length > 0 ? itemsInCategory.filter((item) => item["Jenis Pekerjaan"].toLowerCase().includes(searchTerm)) : itemsInCategory;
    if (itemsToDisplay.length > 0) {
      itemsToDisplay.forEach((item) => {
        const li = document.createElement("li");
        li.textContent = item["Jenis Pekerjaan"];
        li.addEventListener("mousedown", (e) => {
          e.preventDefault();
          searchInput.value = item["Jenis Pekerjaan"];
          hiddenSelect.value = item["Jenis Pekerjaan"];
          suggestionsList.classList.add("hidden");
          autoFillPrices(hiddenSelect);
        });
        suggestionsList.appendChild(li);
      });
      suggestionsList.classList.remove("hidden");
    } else {
      suggestionsList.classList.add("hidden");
    }
  };
  searchInput.addEventListener("input", searchInput._inputHandler);

  searchInput._focusHandler = () => {
    searchInput.dispatchEvent(new Event("input"));
  };
  searchInput.addEventListener("focus", searchInput._focusHandler);

  searchInput._blurHandler = () => {
    setTimeout(() => {
      if (!suggestionsList.contains(document.activeElement)) {
        suggestionsList.classList.add("hidden");
      }
    }, 150);
  };
  searchInput.addEventListener("blur", searchInput._blurHandler);
};

const autoFillPrices = (selectElement) => {
  const row = selectElement.closest("tr");
  const selectedJenisPekerjaan = selectElement.value;
  const currentLingkupPekerjaan = lingkupPekerjaanSelect.value;
  const currentCategory = row.closest(".boq-table-body").dataset.category;

  let selectedItem = null;
  let dataSource = {};

  if (currentLingkupPekerjaan === "Sipil" && categorizedPrices.categorizedSipilPrices) {
    dataSource = categorizedPrices.categorizedSipilPrices;
  } else if (currentLingkupPekerjaan === "ME" && categorizedPrices.categorizedMePrices) {
    dataSource = categorizedPrices.categorizedMePrices;
  }

  if (dataSource[currentCategory]) {
    selectedItem = dataSource[currentCategory].find((item) => item["Jenis Pekerjaan"] === selectedJenisPekerjaan);
  }

  if (selectedItem) {
    row.querySelector(".harga-material").value = selectedItem["Harga Material"];
    row.querySelector(".harga-upah").value = selectedItem["Harga Upah"];
    row.querySelector(".satuan").value = selectedItem["Satuan"];
  } else {
    row.querySelector(".harga-material").value = "";
    row.querySelector(".harga-upah").value = "";
    row.querySelector(".satuan").value = "Ls";
  }
  calculateTotalPrice(row.querySelector(".volume"));
};

const createBoQRow = (category, scope) => {
  const row = document.createElement("tr");
  row.classList.add("boq-item-row");
  row.dataset.category = category;
  row.dataset.scope = scope;
  
  // --- PERUBAHAN: Menambahkan class pada <td> ---
  row.innerHTML = `<td class="col-no"><span class="row-number"></span></td><td class="col-jenis-pekerjaan"><div class="jenis-pekerjaan-wrapper"><input type="text" class="jenis-pekerjaan-search-input" placeholder="Cari Jenis Pekerjaan"><select class="jenis-pekerjaan hidden" name="Jenis_Pekerjaan_Item" required><option value="">-- Pilih --</option></select><ul class="jenis-pekerjaan-suggestions hidden"></ul></div></td><td class="col-satuan"><input type="text" class="satuan" name="Satuan_Item" required readonly /></td><td class="col-volume"><input type="number" class="volume" name="Volume_Item" value="0.00" min="0" step="0.01" /></td><td class="col-harga"><input type="number" class="harga-material" name="Harga_Material_Item" min="0" required readonly /></td><td class="col-harga"><input type="number" class="harga-upah" name="Harga_Upah_Item" min="0" required readonly /></td><td class="col-total"><input type="text" class="total-material" disabled /></td><td class="col-total"><input type="text" class="total-upah" disabled /></td><td class="col-total-harga"><input type="text" class="total-harga" disabled /></td><td class="col-aksi"><button type="button" class="delete-row-btn">Hapus</button></td>`;
  // --- BATAS PERUBAHAN ---

  [row.querySelector(".volume"), row.querySelector(".harga-material"), row.querySelector(".harga-upah")].forEach((input) => {
    input.addEventListener("input", () => calculateTotalPrice(input));
  });
  row.querySelector(".delete-row-btn").addEventListener("click", () => {
    row.remove();
    updateAllRowNumbersAndTotals();
  });
  row.querySelector('.jenis-pekerjaan').addEventListener('change', (e) => autoFillPrices(e.target));
  return row;
};

const updateAllRowNumbersAndTotals = () => {
  document.querySelectorAll(".boq-table-body:not(.hidden)").forEach((tbody) => {
    const rows = tbody.querySelectorAll(".boq-item-row");
    rows.forEach((row, index) => {
      row.querySelector(".row-number").textContent = index + 1;
      calculateTotalPrice(row.querySelector(".volume"));
    });
    calculateSubTotal(tbody);
  });
  calculateGrandTotal();
};

const calculateSubTotal = (tbodyElement) => {
  let subTotal = 0;
  tbodyElement.querySelectorAll(".boq-item-row").forEach((row) => {
    subTotal += parseRupiah(row.querySelector(".total-harga").value);
  });
  const subTotalAmountElement = tbodyElement.closest("table").querySelector(".sub-total-amount");
  if (subTotalAmountElement) subTotalAmountElement.textContent = formatRupiah(subTotal);
};

function calculateTotalPrice(inputElement) {
  const row = inputElement.closest("tr");
  if (!row) return;
  const volume = parseFloat(row.querySelector("input.volume").value) || 0;
  const material = parseFloat(row.querySelector("input.harga-material").value) || 0;
  const upah = parseFloat(row.querySelector("input.harga-upah").value) || 0;
  const totalMaterial = volume * material;
  const totalUpah = volume * upah;
  row.querySelector("input.total-material").value = formatRupiah(totalMaterial);
  row.querySelector("input.total-upah").value = formatRupiah(totalUpah);
  row.querySelector("input.total-harga").value = formatRupiah(totalMaterial + totalUpah);
  calculateSubTotal(row.closest(".boq-table-body"));
  calculateGrandTotal();
}

const calculateGrandTotal = () => {
  let total = 0;
  document.querySelectorAll(".boq-table-body:not(.hidden) .total-harga").forEach((input) => {
    total += parseRupiah(input.value);
  });
  if (grandTotalAmount) grandTotalAmount.textContent = formatRupiah(total);
};

async function initializePage() {
  console.log("Page is being initialized...");
  
  const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzPubDTa7E2gT5HeVLv9edAcn1xaTiT3J4BtAVYqaqiFAvFtp1qovTXpqpm-VuNOxQJ/exec"; 

  form = document.getElementById("form");
  submitButton = document.getElementById("submit-button");
  messageDiv = document.getElementById("message");
  grandTotalAmount = document.getElementById("grand-total-amount");
  lingkupPekerjaanSelect = document.getElementById("lingkup_pekerjaan");
  sipilTablesWrapper = document.getElementById("sipil-tables-wrapper");
  meTablesWrapper = document.getElementById("me-tables-wrapper");
  currentResetButton = form.querySelector("button[type='reset']");
  
  const populateFormWithHistory = (data, message) => {
      console.log("Populating form with rejected data:", data);
      form.reset();
      document.querySelectorAll(".boq-table-body").forEach(tbody => tbody.innerHTML = "");
      sipilTablesWrapper.classList.add("hidden");
      meTablesWrapper.classList.add("hidden");
      
      for (const key in data) {
          if (data.hasOwnProperty(key)) {
              const elementName = key.replace(/_/g, " ");
              const element = document.getElementsByName(elementName)[0];
              if (element) {
                  element.value = (element.type === 'date' && data[key]) ? new Date(data[key]).toISOString().split('T')[0] : data[key];
                  if (key === 'Lingkup_Pekerjaan') {
                      lingkupPekerjaanSelect.dispatchEvent(new Event('change'));
                  }
              }
          }
      }
      
      for (let i = 1; i <= 50; i++) {
          if (data[`Jenis_Pekerjaan_${i}`]) {
              const category = data[`Kategori_Pekerjaan_${i}`];
              const scope = data.Lingkup_Pekerjaan;
              const targetTbody = document.querySelector(`.boq-table-body[data-category="${category}"][data-scope="${scope}"]`);
              if (targetTbody) {
                  const newRow = createBoQRow(category, scope);
                  targetTbody.appendChild(newRow);
                  newRow.querySelector('.jenis-pekerjaan').value = data[`Jenis_Pekerjaan_${i}`];
                  newRow.querySelector('.jenis-pekerjaan-search-input').value = data[`Jenis_Pekerjaan_${i}`];
                  newRow.querySelector('.satuan').value = data[`Satuan_Item_${i}`];
                  newRow.querySelector('.volume').value = data[`Volume_Item_${i}`];
                  newRow.querySelector('.harga-material').value = data[`Harga_Material_Item_${i}`];
                  newRow.querySelector('.harga-upah').value = data[`Harga_Upah_Item_${i}`];
                  populateJenisPekerjaanOptionsForNewRow(newRow);
              }
          }
      }
      updateAllRowNumbersAndTotals();
      messageDiv.innerHTML = message;
      messageDiv.style.display = 'block';
      messageDiv.style.backgroundColor = '#dc3545';
      messageDiv.style.color = 'white';
  };

  const userEmail = sessionStorage.getItem('loggedInUserEmail');
  if (userEmail) {
      try {
          messageDiv.textContent = 'Memeriksa status pengajuan...';
          messageDiv.style.display = 'block';
          const checkUrl = `${APPS_SCRIPT_URL}?action=checkUserStatus&email=${encodeURIComponent(userEmail)}`;
          const response = await fetch(checkUrl);
          const result = await response.json();
          messageDiv.style.display = 'none';
          
          console.log("User submissions response:", result);

          if (result.error) {
              throw new Error(result.error);
          }

          if (result.active_codes) {
              pendingStoreCodes = result.active_codes.pending || [];
              approvedStoreCodes = result.active_codes.approved || [];
          }

          if (result.last_rejected_data) {
              const rejectedStatus = result.last_rejected_data.Status;
              populateFormWithHistory(result.last_rejected_data, `Formulir Anda sebelumnya <strong>${rejectedStatus}</strong>. Silakan periksa, revisi, dan kirim ulang.`);
          }
          
      } catch (error) {
          console.error("Gagal memeriksa status pengajuan:", error);
          messageDiv.textContent = "Gagal memuat status pengajuan terakhir.";
          messageDiv.style.display = 'block';
          messageDiv.style.backgroundColor = '#dc3545';
          messageDiv.style.color = 'white';
      }
  }

  const APPS_SCRIPT_DATA_URL = "https://script.google.com/macros/s/AKfycbx2rtKmaZBb_iRBRL-DOemjVhAp3GaCwsthtwtfdtvdtuO2bRVlmONboB8wE-CZU7Hc/exec"; 
  try {
    const response = await fetch(APPS_SCRIPT_DATA_URL);
    if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
    categorizedPrices = await response.json();
    console.log("Data harga berhasil dimuat.");
  } catch (error) {
    console.error('Error loading price data:', error);
  }

  document.querySelectorAll(".add-row-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const category = button.dataset.category;
      const scope = button.dataset.scope;
      const targetTbody = document.querySelector(`.boq-table-body[data-category="${category}"][data-scope="${scope}"]`);
      if (targetTbody) {
        const newRow = createBoQRow(category, scope);
        targetTbody.appendChild(newRow);
        populateJenisPekerjaanOptionsForNewRow(newRow);
        updateAllRowNumbersAndTotals();
      }
    });
  });

  lingkupPekerjaanSelect.addEventListener("change", (event) => {
    const selectedScope = event.target.value;
    sipilTablesWrapper.classList.toggle("hidden", selectedScope !== 'Sipil');
    meTablesWrapper.classList.toggle("hidden", selectedScope !== 'ME');
    document.querySelectorAll(".boq-table-body").forEach(tbody => {
      if (tbody.dataset.scope !== selectedScope) tbody.innerHTML = "";
    });
    if (selectedScope) {
        document.querySelectorAll(`.boq-table-body[data-scope="${selectedScope}"]`).forEach((tbody) => {
            if (tbody.children.length === 0) {
              const newRow = createBoQRow(tbody.dataset.category, selectedScope);
              tbody.appendChild(newRow);
              populateJenisPekerjaanOptionsForNewRow(newRow);
            }
        });
    }
    updateAllRowNumbersAndTotals();
  });

  currentResetButton.addEventListener("click", () => {
    if (confirm("Apakah Anda yakin ingin mengulang dan mengosongkan semua isian form?")) {
        window.location.reload();
    }
  });

  form.addEventListener("submit", async function (e) {
    e.preventDefault();

    const currentStoreCode = String(document.getElementById('lokasi').value).toUpperCase();

    if (approvedStoreCodes.map(code => String(code).toUpperCase()).includes(currentStoreCode)) {
        messageDiv.textContent = `Error: Kode toko ${currentStoreCode} sudah pernah diajukan dan disetujui.`;
        messageDiv.style.display = "block";
        messageDiv.style.backgroundColor = "#dc3545";
        messageDiv.style.color = "white";
        return;
    }
    
    if (pendingStoreCodes.map(code => String(code).toUpperCase()).includes(currentStoreCode)) {
        messageDiv.textContent = `Error: Kode toko ${currentStoreCode} sudah memiliki pengajuan yang sedang direview.`;
        messageDiv.style.display = "block";
        messageDiv.style.backgroundColor = "#ffc107";
        messageDiv.style.color = "black";
        return;
    }

    messageDiv.textContent = "Mengirim data...";
    messageDiv.style.display = "block";
    messageDiv.style.backgroundColor = '#007bff';
    messageDiv.style.color = 'white';
    submitButton.disabled = true;

    try {
      const formDataToSend = {};
      new FormData(this).forEach((value, key) => formDataToSend[key] = value);
        
      formDataToSend["Email_Pembuat"] = sessionStorage.getItem('loggedInUserEmail') || '';
      formDataToSend["Lokasi"] = currentStoreCode;

      let itemCounter = 0;
      document.querySelectorAll(".boq-table-body:not(.hidden) .boq-item-row").forEach(row => {
        const jenisPekerjaan = row.querySelector(".jenis-pekerjaan").value;
        if (jenisPekerjaan) {
          itemCounter++;
          formDataToSend[`Kategori_Pekerjaan_${itemCounter}`] = row.dataset.category;
          formDataToSend[`Jenis_Pekerjaan_${itemCounter}`] = jenisPekerjaan;
          formDataToSend[`Satuan_Item_${itemCounter}`] = row.querySelector(".satuan").value;
          formDataToSend[`Volume_Item_${itemCounter}`] = parseFloat(row.querySelector(".volume").value) || 0;
          formDataToSend[`Harga_Material_Item_${itemCounter}`] = parseFloat(row.querySelector(".harga-material").value) || 0;
          formDataToSend[`Harga_Upah_Item_${itemCounter}`] = parseFloat(row.querySelector(".harga-upah").value) || 0;
          formDataToSend[`Total_Material_Item_${itemCounter}`] = parseRupiah(row.querySelector(".total-material").value);
          formDataToSend[`Total_Upah_Item_${itemCounter}`] = parseRupiah(row.querySelector(".total-upah").value);
          formDataToSend[`Total_Harga_Item_${itemCounter}`] = parseRupiah(row.querySelector(".total-harga").value);
        }
      });
      formDataToSend["Grand_Total"] = parseRupiah(grandTotalAmount.textContent);

      const response = await fetch(APPS_SCRIPT_URL, {
        method: "POST",
        body: JSON.stringify(formDataToSend),
        headers: { "Content-Type": "text/plain;charset=utf-8" },
      });
      const data = await response.json();

      if (data.status === "success") {
        messageDiv.textContent = data.message || "Data berhasil terkirim! Anda akan diarahkan ke Beranda.";
        messageDiv.style.backgroundColor = "#28a745";
        
        setTimeout(() => {
            window.location.href = '../Homepage/index.html';
        }, 2500);

      } else {
        throw new Error(data.message || "Pengiriman data gagal.");
      }
    } catch (error) {
      console.error("Error submitting form:", error);
      messageDiv.textContent = "Error: " + error.message;
      messageDiv.style.backgroundColor = "#dc3545";
    } finally {
      submitButton.disabled = false;
    }
  });
}

window.addEventListener("pageshow", function(event) {
    initializePage();
});