let parsedAoA = [];
let detectedHeaders = [];

// ---------------- Drag & Drop Helpers -----------------
let dragSrcEl = null;

function handleDragStart(e) {
  dragSrcEl = this;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/html', this.outerHTML);
  this.classList.add('dragging');
}

function handleDragOver(e) {
  e.preventDefault();
  this.classList.add('drag-over');
  e.dataTransfer.dropEffect = 'move';
  return false;
}

function handleDragLeave() {
  this.classList.remove('drag-over');
}

function handleDrop(e) {
  e.stopPropagation();
  if (dragSrcEl !== this) {
    this.parentNode.removeChild(dragSrcEl);
    const dropHTML = e.dataTransfer.getData('text/html');
    this.insertAdjacentHTML('beforebegin', dropHTML);
    const droppedElem = this.previousSibling;
    addDragEventHandlers(droppedElem);
  }
  this.classList.remove('drag-over');
  return false;
}

function handleDragEnd(e) {
  this.classList.remove('dragging');
  document.querySelectorAll('#columnEditor li').forEach(li => {
    li.classList.remove('drag-over');
  });
}

function addDragEventHandlers(li) {
  li.addEventListener('dragstart', handleDragStart);
  li.addEventListener('dragover', handleDragOver);
  li.addEventListener('dragleave', handleDragLeave);
  li.addEventListener('drop', handleDrop);
  li.addEventListener('dragend', handleDragEnd);
}
// -------------------------------------------------------


// This runs every time the taskpane loads or gets re-hydrated
Office.onReady(info => {
  console.log("Office.onReady called");

  if (info.host === Office.HostType.Excel) {
    console.log("Host is Excel. Wiring up event handlers…");

    const importBtn = document.getElementById("importBtn");
    const filePicker = document.getElementById("filePicker");
    const editParamsBtn = document.getElementById("editParams");
    const closeModalBtn = document.getElementById("closeModal");
    const fillDataBtn = document.getElementById("fillDataBtn");
    const editorList = document.getElementById("columnEditor");

    // ---- 1) REMOVE existing handlers if present ----
    console.log("Removing any existing handlers (safe if not present)...");
    importBtn.removeEventListener("click", importBtnClickHandler);
    filePicker.removeEventListener("change", handleFilePick);
    editParamsBtn.removeEventListener("click", openEditModal);
    closeModalBtn.removeEventListener("click", closeEditModal);
    fillDataBtn.removeEventListener("click", fillDataIntoExcel);
    editorList.removeEventListener("click", columnEditorHandler);

    // ---- 2) ADD handlers ----
    console.log("Attaching handlers now.");
    importBtn.addEventListener("click", importBtnClickHandler);
    filePicker.addEventListener("change", handleFilePick);
    editParamsBtn.addEventListener("click", openEditModal);
    closeModalBtn.addEventListener("click", closeEditModal);
    fillDataBtn.addEventListener("click", fillDataIntoExcel);
    editorList.addEventListener("click", columnEditorHandler);
  }
});

// Separate named handler for import button
function importBtnClickHandler(e) {
  console.log("Import button clicked");
  e.preventDefault();
  const filePicker = document.getElementById("filePicker");
  // Reset and open picker
  filePicker.value = "";
  filePicker.click();
}

async function handleFilePick(evt) {
  console.log("File picker change event triggered");

  const statusEl = document.getElementById("fileStatus");
  const listEl = document.getElementById("paramList");

  // Clear previous results  
  listEl.innerHTML = "";
  detectedHeaders = [];
  parsedAoA = [];

  const file = evt.target.files?.[0];
  if (!file) {
    statusEl.textContent = "No file selected.";
    console.log("No file selected");
    return;
  }

  statusEl.textContent = `Reading "${file.name}"...`;
  console.log(`Reading file ${file.name}…`);

  try {
    const arrayBuf = await file.arrayBuffer();
    const wb = XLSX.read(arrayBuf, { type: "array" });

    if (!wb.SheetNames.length) {
      throw new Error("No sheets found in file.");
    }

    const wsName = wb.SheetNames[0];
    const ws = wb.Sheets[wsName];
    parsedAoA = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });

    if (!parsedAoA.length) {
      throw new Error("The selected sheet is empty.");
    }

    detectedHeaders = parsedAoA[0].map(h =>
      h == null ? "" : String(h)
    );

    console.log("Detected headers:", detectedHeaders);

    document.getElementById("paramTitle").textContent =
      `${detectedHeaders.length} Parameters`;

    detectedHeaders.forEach((name, i) => {
      const li = document.createElement("li");
      li.textContent = `${columnLetter(i)} ${name || "(Unnamed Column)"}`;
      listEl.appendChild(li);
    });

    statusEl.textContent =
      `Loaded "${wsName}": ${parsedAoA.length - 1} data rows, ${detectedHeaders.length} columns.`;
    document.getElementById("editParams").disabled = false;

  } catch (err) {
    console.error(err);
    statusEl.textContent = `Error: ${err.message}`;
  }
}

function openEditModal() {
  console.log("Opening edit modal");
  const editorList = document.getElementById("columnEditor");
  editorList.innerHTML = "";

  detectedHeaders.forEach((header, index) => {
    const li = document.createElement("li");
    li.dataset.index = index;
    li.draggable = true;
    li.classList.add("draggable-row");
    li.innerHTML = `
       <span class="handle">⋮⋮</span>
       <input type="checkbox" checked />
       <span class="label">${header || "(Unnamed Column)"}</span>
    `;
    addDragEventHandlers(li); // <-- add DRAG handlers
    editorList.appendChild(li);
  });

  document.getElementById("editModal").classList.add("show");
}



function closeEditModal() {
  console.log("Closing edit modal");
  document.getElementById("editModal").classList.remove("show");
}

function columnEditorHandler(e) {
  if (e.target.tagName === "BUTTON") {
    const index = parseInt(e.target.closest("li").dataset.index, 10);
    if (e.target.dataset.action === "up") {
      console.log(`Move up column ${index}`);
      moveColumn(index, -1);
    }
    if (e.target.dataset.action === "down") {
      console.log(`Move down column ${index}`);
      moveColumn(index, 1);
    }
    if (e.target.dataset.action === "delete") {
      console.log(`Delete column ${index}`);
      deleteColumn(index);
    }
  }
}

function moveColumn(index, direction) {
  const newIndex = index + direction;
  if (newIndex < 0 || newIndex >= detectedHeaders.length) return;

  [detectedHeaders[index], detectedHeaders[newIndex]] =
    [detectedHeaders[newIndex], detectedHeaders[index]];

  parsedAoA.forEach(row => {
    [row[index], row[newIndex]] = [row[newIndex], row[index]];
  });

  openEditModal();
}

function deleteColumn(index) {
  detectedHeaders.splice(index, 1);
  parsedAoA.forEach(row => row.splice(index, 1));
  openEditModal();
}

async function fillDataIntoExcel() {
  console.log("Save (fill) clicked");

  // Build array of ACTIVE columns
  const editorItems = [...document.querySelectorAll("#columnEditor li")];
  const activeIndexes = [];
  editorItems.forEach((li, i) => {
    const checkbox = li.querySelector("input[type=checkbox]");
    if (checkbox.checked) {
      activeIndexes.push(parseInt(li.dataset.index, 10));
    }
  });

  // Build a filtered version of parsedAoA using only the activeIndexes
  const filteredAoA = parsedAoA.map(row => {
    return activeIndexes.map(i => row[i]);
  });

  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRangeByIndexes(
      0,
      0,
      filteredAoA.length,
      activeIndexes.length
    );
    range.values = filteredAoA;
    await context.sync();
  });

  closeEditModal();
}


// Helper to get Excel-style column letters
function columnLetter(n) {
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}
