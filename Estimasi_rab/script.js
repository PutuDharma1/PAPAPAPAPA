// Estimasi_rab/script.js

// --- Global Variable Declarations (initialized in window.onload) ---
let form;
let submitButton;
let messageDiv;
let grandTotalAmount;
let lingkupPekerjaanSelect;
let sipilTablesWrapper;
let meTablesWrapper;
let currentResetButton; // Reference to the reset button

// Data storage for prices, loaded from Google Apps Script
let categorizedPrices = {};

// --- Helper Functions ---

/**
 * Formats a number into Indonesian Rupiah currency string.
 * @param {number} number - The number to format.
 * @returns {string} The formatted Rupiah string (e.g., "Rp 1.000.000").
 */
const formatRupiah = (number) => {
  return new Intl.NumberFormat("id-ID", {
    style: "currency",
    currency: "IDR",
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(number);
};

/**
 * Parses a Rupiah formatted string back into a number.
 * Handles common Indonesian Rupiah formatting (Rp, dots for thousands, commas for decimals).
 * @param {string} formattedString - The Rupiah string to parse.
 * @returns {number} The parsed number (e.g., 1000000).
 */
const parseRupiah = (formattedString) => {
  const cleanedString = formattedString
    .replace(/Rp\s?|\./g, "") // Remove "Rp " and dots
    .replace(/,/g, "."); // Replace comma with dot for decimal (though your format doesn't use it for decimals)
  return parseFloat(cleanedString) || 0; // Convert to float, default to 0 if invalid
};

/**
 * Populates "Jenis Pekerjaan" options for a new row, including search functionality.
 * Attaches event listeners for the search input and hidden select.
 * @param {HTMLElement} rowElement - The table row (<tr>) element to populate.
 */
const populateJenisPekerjaanOptionsForNewRow = (rowElement) => {
  const category = rowElement.dataset.category;
  const scope = rowElement.dataset.scope;
  const searchInput = rowElement.querySelector(".jenis-pekerjaan-search-input");
  const hiddenSelect = rowElement.querySelector(".jenis-pekerjaan");
  const suggestionsList = rowElement.querySelector(".jenis-pekerjaan-suggestions");

  // Basic validation to ensure elements exist
  if (!searchInput || !hiddenSelect || !suggestionsList) {
    console.error("Error: Required elements for Jenis Pekerjaan dropdown not found in row:", rowElement);
    return;
  }

  let dataSource = {};
  if (scope === "Sipil" && categorizedPrices.categorizedSipilPrices) {
    dataSource = categorizedPrices.categorizedSipilPrices;
  } else if (scope === "ME" && categorizedPrices.categorizedMePrices) {
    dataSource = categorizedPrices.categorizedMePrices;
  } else {
    console.warn(`No data source found for scope: ${scope}`);
    return;
  }

  const itemsInCategory = dataSource[category] || [];

  // Clear existing options and add a default one
  hiddenSelect.innerHTML = '<option value="">-- Pilih Jenis Pekerjaan --</option>';

  // Populate the hidden select with options
  itemsInCategory.forEach((item) => {
    const option = document.createElement("option");
    option.value = item["Jenis Pekerjaan"];
    option.textContent = item["Jenis Pekerjaan"];
    hiddenSelect.appendChild(option);
  });

  // Remove existing listeners to prevent duplicates (important for dynamically added rows)
  const removeListeners = (inputEl, selectEl, suggestionsEl) => {
    if (inputEl._inputHandler) inputEl.removeEventListener("input", inputEl._inputHandler);
    if (inputEl._focusHandler) inputEl.removeEventListener("focus", inputEl._focusHandler);
    if (inputEl._blurHandler) inputEl.removeEventListener("blur", inputEl._blurHandler);
    if (selectEl._changeHandler) selectEl.removeEventListener("change", selectEl._changeHandler);
  };
  removeListeners(searchInput, hiddenSelect, suggestionsList); // Call before adding new ones

  // --- Event Listeners for Search Input and Hidden Select ---

  searchInput._inputHandler = () => {
    const searchTerm = searchInput.value.toLowerCase();
    suggestionsList.innerHTML = ""; // Clear previous suggestions

    const itemsToDisplay =
      searchTerm.length > 0
        ? itemsInCategory.filter((item) =>
            item["Jenis Pekerjaan"].toLowerCase().includes(searchTerm)
          )
        : itemsInCategory; // Show all if search term is empty

    if (itemsToDisplay.length > 0) {
      itemsToDisplay.forEach((item) => {
        const li = document.createElement("li");
        li.textContent = item["Jenis Pekerjaan"];
        li.classList.add("p-2", "hover:bg-gray-200", "cursor-pointer"); // Tailwind classes for styling suggestions
        li.addEventListener("mousedown", (e) => {
          e.preventDefault(); // Prevent blur event from immediately closing suggestions
          searchInput.value = item["Jenis Pekerjaan"];
          hiddenSelect.value = item["Jenis Pekerjaan"]; // Update hidden select value
          suggestionsList.classList.add("hidden"); // Hide suggestions
          autoFillPrices(hiddenSelect); // Trigger autofill based on selection
        });
        suggestionsList.appendChild(li);
      });
      suggestionsList.classList.remove("hidden"); // Show suggestions list
    } else {
      suggestionsList.classList.add("hidden"); // Hide if no matches
    }
  };
  searchInput.addEventListener("input", searchInput._inputHandler);

  searchInput._focusHandler = () => {
    // Show suggestions when input is focused, triggering initial input event
    searchInput.dispatchEvent(new Event("input"));
  };
  searchInput.addEventListener("focus", searchInput._focusHandler);

  searchInput._blurHandler = () => {
    // Hide suggestions after a short delay to allow click on suggestions
    setTimeout(() => {
      // Check if focus moved to a suggestion item or still on the input
      if (
        !suggestionsList.contains(document.activeElement) &&
        document.activeElement !== searchInput
      ) {
        suggestionsList.classList.add("hidden");
      }
    }, 100);
  };
  searchInput.addEventListener("blur", searchInput._blurHandler);

  hiddenSelect._changeHandler = () => {
    autoFillPrices(hiddenSelect); // Trigger autofill when hidden select changes (e.g., from initial value)
    searchInput.value = hiddenSelect.value; // Sync search input with hidden select
  };
  hiddenSelect.addEventListener("change", hiddenSelect._changeHandler);

  // If a value is already set (e.g., from a pre-filled form or initial row setup)
  if (hiddenSelect.value) {
    searchInput.value = hiddenSelect.value;
    autoFillPrices(hiddenSelect);
  }
};

/**
 * Automatically fills material and labor prices based on the selected "Jenis Pekerjaan".
 * @param {HTMLSelectElement} selectElement - The select element for "Jenis Pekerjaan".
 */
const autoFillPrices = (selectElement) => {
  const row = selectElement.closest("tr");
  const selectedJenisPekerjaan = selectElement.value;
  const currentLingkupPekerjaan = lingkupPekerjaanSelect.value; // Get the main lingkup selected
  const currentCategory = row.closest(".boq-table-body").dataset.category;

  let selectedItem = null;
  let dataSource = {};

  if (currentLingkupPekerjaan === "Sipil" && categorizedPrices.categorizedSipilPrices) {
    dataSource = categorizedPrices.categorizedSipilPrices;
  } else if (currentLingkupPekerjaan === "ME" && categorizedPrices.categorizedMePrices) {
    dataSource = categorizedPrices.categorizedMePrices;
  }

  // Find the selected item within its category
  if (dataSource[currentCategory]) {
    selectedItem = dataSource[currentCategory].find(
      (item) => item["Jenis Pekerjaan"] === selectedJenisPekerjaan
    );
  }

  if (selectedItem) {
    row.querySelector(".harga-material").value = selectedItem["Harga Material"];
    row.querySelector(".harga-upah").value = selectedItem["Harga Upah"];
    row.querySelector(".satuan").value = selectedItem["Satuan"]; // Auto-fill satuan
  } else {
    // Clear values if no match found or option is empty
    row.querySelector(".harga-material").value = "";
    row.querySelector(".harga-upah").value = "";
    row.querySelector(".satuan").value = "Ls"; // Default to Ls if nothing is found
  }
  calculateTotalPrice(row.querySelector(".volume")); // Recalculate totals for this row
};

/**
 * Creates and returns a new BOQ table row element.
 * @param {string} category - The category of the work (e.g., "PEKERJAAN PERSIAPAN").
 * @param {string} scope - The scope of work (e.g., "Sipil" or "ME").
 * @returns {HTMLElement} The created table row.
 */
const createBoQRow = (category, scope) => {
  const row = document.createElement("tr");
  row.classList.add("boq-item-row");
  row.dataset.category = category;
  row.dataset.scope = scope; // Store scope on the row for easier lookup

  row.innerHTML = `
    <td><span class="row-number"></span></td>
    <td>
      <div class="jenis-pekerjaan-wrapper relative">
        <input type="text" class="jenis-pekerjaan-search-input p-2 border rounded-md w-full" placeholder="Cari Jenis Pekerjaan">
        <select class="jenis-pekerjaan hidden" name="Jenis_Pekerjaan_Item" required> <option value="">-- Pilih Jenis Pekerjaan --</option>
        </select>
        <ul class="jenis-pekerjaan-suggestions absolute bg-white border border-gray-300 w-full max-h-48 overflow-y-auto z-10 hidden mt-1 rounded-md shadow-lg">
        </ul>
      </div>
    </td>
    <td>
      <select class="satuan p-2 border rounded-md" name="Satuan_Item" required>
        <option value="Ls">Ls</option>
        <option value="M1">M1</option>
        <option value="M3">M3</option>
        <option value="Btg">Btg</option>
        <option value="M2">M2</option>
        <option value="Bh">Bh</option>
        <option value="Unit">Unit</option>
        <option value="Kg">Kg</option>
        <option value="sel">sel</option>
        <option value="ttk">ttk</option>
        <option value="m">m</option>
      </select>
    </td>
    <td>
      <input type="number" class="volume p-2 border rounded-md" name="Volume_Item" value="1.00" min="0" step="0.01" />
    </td>
    <td>
      <input type="number" class="harga-material p-2 border rounded-md" name="Harga_Material_Item" min="0" required />
    </td>
    <td>
      <input type="number" class="harga-upah p-2 border rounded-md" name="Harga_Upah_Item" min="0" required />
    </td>
    <td>
      <input type="text" class="total-material p-2 border rounded-md" disabled />
    </td>
    <td>
      <input type="text" class="total-upah p-2 border rounded-md" disabled />
    </td>
    <td>
      <input type="text" class="total-harga p-2 border rounded-md" disabled />
    </td>
    <td>
      <button type="button" class="delete-row-btn bg-red-500 text-white p-2 rounded-md hover:bg-red-600 transition-colors">Hapus</button>
    </td>
  `;

  // Attach event listeners to the inputs within the new row
  const volumeInput = row.querySelector(".volume");
  const materialInput = row.querySelector(".harga-material");
  const upahInput = row.querySelector(".harga-upah");
  const deleteButton = row.querySelector(".delete-row-btn");

  [volumeInput, materialInput, upahInput].forEach((input) => {
    input.addEventListener("input", () => {
      calculateTotalPrice(input);
      calculateGrandTotal(); // Update grand total whenever a row's total changes
    });
  });

  // Attach event listener for the delete button
  deleteButton.addEventListener("click", () => {
    row.remove(); // Remove the row from the DOM
    updateAllRowNumbersAndTotals(); // Recalculate and update all numbers and totals
  });

  return row;
};

/**
 * Updates row numbers, recalculates total for each row, and updates sub-totals and grand total.
 * This should be called after adding/removing rows or after a reset.
 */
const updateAllRowNumbersAndTotals = () => {
  // Select all visible table bodies
  const visibleTableBodies = document.querySelectorAll(".boq-table-body:not(.hidden)");
  visibleTableBodies.forEach((tbody) => {
    const rows = tbody.querySelectorAll(".boq-item-row");
    rows.forEach((row, index) => {
      row.querySelector(".row-number").textContent = index + 1; // Update row number
      const hiddenSelect = row.querySelector(".jenis-pekerjaan");
      // If a job type is selected, re-autofill prices (in case data changed, or initial fill)
      if (hiddenSelect && hiddenSelect.value) {
        autoFillPrices(hiddenSelect);
      } else {
        // Otherwise, just calculate based on existing inputs
        calculateTotalPrice(row.querySelector(".volume"));
      }
    });
    calculateSubTotal(tbody); // Recalculate subtotal for each table body
  });
  // Finally, update the grand total
  if (grandTotalAmount) {
    calculateGrandTotal();
  }
};

/**
 * Calculates the subtotal for a specific table body.
 * @param {HTMLElement} tbodyElement - The tbody element for which to calculate the subtotal.
 */
const calculateSubTotal = (tbodyElement) => {
  let subTotal = 0;
  const rows = tbodyElement.querySelectorAll(".boq-item-row");
  rows.forEach((row) => {
    const totalHargaInput = row.querySelector(".total-harga");
    if (totalHargaInput) {
      subTotal += parseRupiah(totalHargaInput.value);
    }
  });
  const subTotalAmountElement =
    tbodyElement.closest("table").querySelector(".sub-total-amount");
  if (subTotalAmountElement) {
    subTotalAmountElement.textContent = formatRupiah(subTotal);
  }
};

/**
 * Calculates the total material, total labor, and total price for a single row.
 * @param {HTMLInputElement} inputElement - An input element (volume, material, or upah) from the row that triggered the calculation.
 */
function calculateTotalPrice(inputElement) {
  const row = inputElement.closest("tr");
  if (!row) return; // Exit if row is not found

  const volume = parseFloat(row.querySelector("input.volume").value) || 0;
  const material = parseFloat(row.querySelector("input.harga-material").value) || 0;
  const upah = parseFloat(row.querySelector("input.harga-upah").value) || 0;

  const totalMaterial = volume * material;
  const totalUpah = volume * upah;
  const totalHarga = totalMaterial + totalUpah;

  row.querySelector("input.total-material").value = formatRupiah(totalMaterial);
  row.querySelector("input.total-upah").value = formatRupiah(totalUpah);
  row.querySelector("input.total-harga").value = formatRupiah(totalHarga);

  // Also update the subtotal of the parent table and grand total
  calculateSubTotal(row.closest(".boq-table-body"));
  calculateGrandTotal();
}

/**
 * Calculates the grand total from all visible table rows across all tables.
 */
const calculateGrandTotal = () => {
  let total = 0;
  // Select all table bodies that are currently visible
  const visibleTableBodies = document.querySelectorAll(
    ".boq-table-body:not(.hidden)"
  );
  visibleTableBodies.forEach((tbody) => {
    const rows = tbody.querySelectorAll(".boq-item-row");
    rows.forEach((row) => {
      const totalHargaInput = row.querySelector(".total-harga");
      if (totalHargaInput) {
        total += parseRupiah(totalHargaInput.value);
      }
    });
  });
  if (grandTotalAmount) {
    grandTotalAmount.textContent = formatRupiah(total);
  }
};

// --- Main Script Execution (Waits for DOM to be fully loaded) ---
window.addEventListener("load", async () => {
  console.log("Window loaded: Initializing script...");

  // Load data from Google Spreadsheet via Apps Script
  const APPS_SCRIPT_DATA_URL = "https://script.google.com/macros/s/AKfycbx2rtKmaZBb_iRBRL-DOemjVhAp3GaCwsthtwtfdtvdtuO2bRVlmONboB8wE-CZU7Hc/exec"; 

  try {
    const response = await fetch(APPS_SCRIPT_DATA_URL);
    if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
    }
    categorizedPrices = await response.json();
    console.log("Data loaded successfully from Apps Script:", categorizedPrices);
  } catch (error) {
    console.error('Error loading data from Google Apps Script:', error);
    messageDiv = document.getElementById("message");
    if (messageDiv) {
      messageDiv.textContent = "Error loading price data. Please check console for details.";
      messageDiv.style.backgroundColor = "#ce1e10";
      messageDiv.style.color = "white";
    }
    submitButton = document.getElementById("submit-button");
    if (submitButton) {
      submitButton.disabled = true;
    }
    return;
  }

  // Get references to DOM elements 
  form = document.getElementById("form");
  submitButton = document.getElementById("submit-button");
  messageDiv = document.getElementById("message");
  grandTotalAmount = document.getElementById("grand-total-amount");
  lingkupPekerjaanSelect = document.getElementById("lingkup_pekerjaan");
  sipilTablesWrapper = document.getElementById("sipil-tables-wrapper");
  meTablesWrapper = document.getElementById("me-tables-wrapper");
  currentResetButton = form.querySelector("button[type='reset']");

  // Initial state: hide both wrappers 
  sipilTablesWrapper.classList.add("hidden");
  meTablesWrapper.classList.add("hidden");

  // --- Event Listeners ---

  // Handle "Add Row" buttons 
  document.querySelectorAll(".add-row-btn").forEach((button) => {
    button.addEventListener("click", () => {
      const category = button.dataset.category;
      const scope = button.dataset.scope;
      const targetTbody = document.querySelector(
        `.boq-table-body[data-category="${category}"][data-scope="${scope}"]`
      );

      if (targetTbody) {
        const newRow = createBoQRow(category, scope);
        targetTbody.appendChild(newRow);
        populateJenisPekerjaanOptionsForNewRow(newRow);
        updateAllRowNumbersAndTotals();
      } else {
        console.error(`Target tbody not found for category: ${category}, scope: ${scope}`);
      }
    });
  });

  // Handle "Lingkup Pekerjaan" select change 
  lingkupPekerjaanSelect.addEventListener("change", (event) => {
    const selectedScope = event.target.value;
    console.log("Lingkup Pekerjaan changed to:", selectedScope);

    document.querySelectorAll(".boq-table-body").forEach((tbody) => {
      tbody.innerHTML = "";
      calculateSubTotal(tbody);
    });
    sipilTablesWrapper.classList.add("hidden");
    meTablesWrapper.classList.add("hidden");
    calculateGrandTotal();

    let activeWrapper = null;
    if (selectedScope === "Sipil") {
      sipilTablesWrapper.classList.remove("hidden");
      activeWrapper = sipilTablesWrapper;
    } else if (selectedScope === "ME") {
      meTablesWrapper.classList.remove("hidden");
      activeWrapper = meTablesWrapper;
    }

    if (activeWrapper) {
      activeWrapper.querySelectorAll(".boq-table-body").forEach((tbody) => {
        const category = tbody.dataset.category;
        const scope = tbody.dataset.scope;
        const newRow = createBoQRow(category, scope);
        tbody.appendChild(newRow);
        populateJenisPekerjaanOptionsForNewRow(newRow);
      });
    }
    updateAllRowNumbersAndTotals();
  });

  // Handle form reset button 
  currentResetButton.addEventListener("click", function () {
    console.log("Form reset initiated.");
    form.reset();
    messageDiv.style.display = "none";

    document.querySelectorAll(".boq-table-body").forEach((tbody) => {
      tbody.innerHTML = "";
      calculateSubTotal(tbody);
    });

    sipilTablesWrapper.classList.add("hidden");
    meTablesWrapper.classList.add("hidden");
    calculateGrandTotal();

    const currentLingkupAfterReset = lingkupPekerjaanSelect.value;
    if (currentLingkupAfterReset) {
      if (currentLingkupAfterReset === "Sipil") {
        sipilTablesWrapper.classList.remove("hidden");
      } else if (currentLingkupAfterReset === "ME") {
        meTablesWrapper.classList.remove("hidden");
      }

      document
        .querySelectorAll(
          `.boq-table-body[data-scope="${currentLingkupAfterReset}"]`
        )
        .forEach((tbody) => {
          const category = tbody.dataset.category;
          const newRow = createBoQRow(category, currentLingkupAfterReset);
          tbody.appendChild(newRow);
          populateJenisPekerjaanOptionsForNewRow(newRow);
        });
    }
    updateAllRowNumbersAndTotals();
  });

  // --- Form Submission Handler --- 
  form.addEventListener("submit", async function (e) {
    e.preventDefault();

    messageDiv.textContent = "Mengirim data...";
    messageDiv.style.display = "block";
    messageDiv.style.backgroundColor = "beige";
    messageDiv.style.color = "black";
    submitButton.disabled = true;
    submitButton.style.opacity = "0.7";

    try {
      const formData = new FormData(this);
      const formDataToSend = {};

      formDataToSend["Timestamp"] = new Date().toLocaleString("en-GB", { timeZone: 'Asia/Jakarta' });
      formDataToSend["Lokasi"] = formData.get("Lokasi") || "";
      formDataToSend["Proyek"] = formData.get("Proyek") || "";
      formDataToSend["Cabang"] = formData.get("Cabang") || "";
      formDataToSend["Lingkup_Pekerjaan"] = formData.get("Lingkup Pekerjaan") || "";
      formDataToSend["Luas_Bangunan"] = parseFloat(formData.get("Luas Bangunan")) || 0;
      formDataToSend["Luas_Terbangunan"] = parseFloat(formData.get("Luas Terbangunan")) || 0;
      formDataToSend["Luas_Area_Terbuka_Area_Parkir"] = parseFloat(formData.get("Luas Area Terbuka / Area Parkir")) || 0;
      formDataToSend["Luas_Area_Sales"] = parseFloat(formData.get("Luas Area Sales")) || 0;
      formDataToSend["Luas_Gudang"] = parseFloat(formData.get("Luas Gudang")) || 0;
      formDataToSend["Tanggal"] = formData.get("Tanggal") || "";
      formDataToSend["Waktu_Pelaksanaan"] = formData.get("Waktu Pelaksanaan") || "";

      const allItemRows = document.querySelectorAll(
        ".boq-table-body:not(.hidden) .boq-item-row"
      );
      let itemCounter = 0;

      allItemRows.forEach((row) => {
        const selectedJenisPekerjaanValue = row.querySelector(".jenis-pekerjaan").value;

        if (selectedJenisPekerjaanValue) {
          itemCounter++;
          
          // --- PERBAIKAN KRUSIAL ADA DI SINI ---
          const category = row.closest('.boq-table-body').dataset.category;
          
          formDataToSend[`No_Item_${itemCounter}`] = itemCounter;
          formDataToSend[`Kategori_Pekerjaan_${itemCounter}`] = category;
          formDataToSend[`Jenis_Pekerjaan_${itemCounter}`] = selectedJenisPekerjaanValue;
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

      console.log("Data being sent to Google Apps Script:", formDataToSend);
      
      const scriptURL = "https://script.google.com/macros/s/AKfycbzPubDTa7E2gT5HeVLv9edAcn1xaTiT3J4BtAVYqaqiFAvFtp1qovTXpqpm-VuNOxQJ/exec";

      const response = await fetch(scriptURL, {
        redirect: "follow",
        method: "POST",
        body: JSON.stringify(formDataToSend),
        headers: { "Content-Type": "text/plain;charset=utf-8" },
      });

      const data = await response.json();
      console.log("Response from Apps Script:", data);

      if (data.status === "success") {
        messageDiv.textContent = data.message || "Data berhasil terkirim!";
        messageDiv.style.backgroundColor = "#28a745";
        messageDiv.style.color = "white";
        currentResetButton.click(); // Memicu reset untuk membersihkan form
      } else {
        throw new Error(data.message || "Pengiriman data gagal.");
      }
    } catch (error) {
      console.error("Error submitting form:", error);
      messageDiv.textContent = "Error: " + error.message;
      messageDiv.style.backgroundColor = "#ce1e10";
      messageDiv.style.color = "white";
    } finally {
      submitButton.disabled = false;
      submitButton.style.opacity = "1";
      setTimeout(() => {
        messageDiv.textContent = "";
        messageDiv.style.display = "none";
      }, 4000);
    }
  });

  if (lingkupPekerjaanSelect.value) {
    lingkupPekerjaanSelect.dispatchEvent(new Event("change"));
  }
});