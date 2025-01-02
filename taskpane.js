/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

class AutoSave {
  constructor(writeSave) {
    this.enabled = false;
    this.writeSave = writeSave;
    this.queue = new Map();
    this.debounceTime = 300; // 0.3 second debounce
  }

  handleInput(inputId, value, onStatusUpdate) {
    console.log("Handling more input");
    // only initiate save if enabled
    if (!this.enabled) {
      return;
    }

    onStatusUpdate(inputId, "pending"); // Update status to pending

    if (this.queue.has(inputId)) {
      clearTimeout(this.queue.get(inputId).timeout);
    }

    const timeout = setTimeout(() => {
      this.saveInput(inputId, value, onStatusUpdate);
    }, this.debounceTime);

    this.queue.set(inputId, { value, timeout });
  }

  async saveInput(inputId, value, onStatusUpdate) {
    this.queue.delete(inputId);

    onStatusUpdate(inputId, "pending"); // Update status to pending

    // ISSUE: not catching errors properly
    const [, err] = await this.writeSave(inputId, value);
    console.log(err);

    if (err) {
      console.log(err.type);
      if (err.type === "ChangeLogNotFound") return;
      onStatusUpdate(inputId, "error"); // Update status to error
      console.error(`Error saving ${inputId}:`, err);
      return;
    }

    // Success! Update status to saved
    onStatusUpdate(inputId, "saved");
  }

  clearQueue() {
    // Iterate through all queued items and clear their timeouts
    for (const [, queueItem] of this.queue) {
      clearTimeout(queueItem.timeout);
    }

    // Clear the entire queue
    this.queue.clear();
  }
  clearFromQueue(inputId) {
    if (this.queue.has(inputId)) {
      clearTimeout(this.queue.get(inputId).timeout);
      this.queue.delete(inputId);
    }
  }
}

const writeDescriptionUpdate = async (inputId, newVal) => {
  console.log("Executing save");

  // called at completion of auto-save
  // overwite only the description
  const [result, error] = await pushEntryToChangeLog(inputId, getInitials(), [[null, null, null, newVal]], true);

  if (error && error.type === "ChangeLogNotFound") {
    await updateCardContainer();
  }
  return [result, error];
};

// Initialize an empty array to store captured ranges
window.sharedState = {
  capturedRanges: [],
  autoSave: new AutoSave(writeDescriptionUpdate),
};

function getCardState(entryId) {
  return window.sharedState.capturedRanges.find((card) => card.id === entryId)?.state;
}
function setCardState(entryId, state) {
  let foundCard = window.sharedState.capturedRanges.find((card) => card.id === entryId);
  if (foundCard) {
    foundCard.state = state;
  } else {
    throw new Error("Unable to set card, as no card with that ID.");
  }
}
function getCardEntryCells(entryId) {
  const data = window.sharedState.capturedRanges.find((card) => card.id === entryId);
  const initials = getInitials();

  // Create a horizontal array with initials and address
  return constructEntryCells(data, initials);
}

// document.addEventListener('DOMContentLoaded', () => {
//     // Your code here
//     const editableDiv = document.getElementById('your-editable-div-id');

//     editableDiv.addEventListener('keydown', handle_enter);

//   });

// ERROR HANDLING

// Global error handler - doesn't work
// window.onerror = function (message, source, lineno, colno, error) {
//   console.error("Uncaught error:", error);
//   showErrorMessage("Uncaught error:" + error);
//   // Log error or display to user
//   return true; // Prevents the firing of the default event handler
// };

const tryCatch = async (callback, showErr = true) => {
  try {
    // await required otherwise throws global error
    const result = await callback();
    return [result, null];
  } catch (error) {
    console.error(error);
    if (showErr) showPopup(error);
    return [null, error];
  }
};

// Cria button

const openCria = () =>
  tryCatch(async () => {
    window.open("https://cria.fiecon.com/", "_blank");
  });

// SAVE INITIALS

async function loadSavedInitials() {
  const initialsInput = document.getElementById("initialsInput");

  // Set the initial value of the input using the load function
  let [result, error] = await getFromLocalStorage("pwrups_user_initials");
  if (!error) initialsInput.value = result;

  // Add event listener for the 'blur' event (when the input loses focus)
  initialsInput.addEventListener("blur", () => setInLocalStorage("pwrups_user_initials", initialsInput.value));
}

async function loadSavedFillColour() {
  const colourPicker = document.getElementById("fill-colour-input");
  let selectedColor;

  let [result, error] = await getFromLocalStorage("pwrups_fill_col");
  if (!error && result) {
    selectedColor = result;
  } else {
    selectedColor = getRandomVibrantColor();
    setInLocalStorage("pwrups_fill_col", selectedColor);
  }
  console.log(selectedColor);
  colourPicker.value = selectedColor;
  colourPicker.addEventListener("input", () => setInLocalStorage("pwrups_fill_col", colourPicker.value));
}

const setInLocalStorage = async (key, value) =>
  tryCatch(async () => {
    const myPartitionKey = Office.context.partitionKey;

    // Check if local storage is partitioned.
    // If so, use the partition to ensure the data is only accessible by your add-in.
    if (myPartitionKey) {
      localStorage.setItem(myPartitionKey + key, value);
    } else {
      localStorage.setItem(key, value);
    }
  });

const getFromLocalStorage = async (key) => {
  return await tryCatch(async () => {
    const myPartitionKey = Office.context.partitionKey;

    // Check if local storage is partitioned.
    if (myPartitionKey) {
      return localStorage.getItem(myPartitionKey + key);
    } else {
      return localStorage.getItem(key);
    }
  });
};

Office.onReady(async (info) => {
  // load initials from user settings
  loadSavedInitials();
  loadSavedFillColour();

  // try and enable autosave (with validation checks)
  const autosaveToggle = document.getElementById("autosaveToggle");
  autosaveToggle.checked = true;
  handleAutosaveToggleChange(autosaveToggle, true);

  console.log("Ready");
});

// CHANGE LOGGER

const captureAddress = async (context = undefined) => {
  // note can probably remove assignment on line below
  let [, error] = await tryCatch(async () => {
    const rangeLimit = window.sharedState.autoSave.enabled ? 20 : 5;

    if (window.sharedState.capturedRanges.length >= rangeLimit) {
      throw Error(
        "Unable to save more than 5 addresses at once. Unload the current addresses to the change log before continuing."
      );
    }

    await Excel.run(async (context) => {
      // get selected range and current worksheet
      const range = context.workbook.getSelectedRange();

      // load and sync required data
      range.load("address, worksheet, worksheet/name");
      await context.sync();

      const ws = range.worksheet;

      // construct data
      const sheet = ws.name;
      const address = range.address.split("!").pop();
      const fullAddress = range.address;
      const description = getDescription();

      // construct ID
      const initials = getInitials();
      const id = generateID(initials, address);

      // autosave for easier ref
      const autoSave = window.sharedState.autoSave;

      const capturedData = {
        id,
        sheet,
        address,
        fullAddress,
        description,
        state: autoSave.enabled ? "none" : "arrow",
      };
      window.sharedState.capturedRanges.push(capturedData);

      await updateCardContainer(true);

      if (autoSave.enabled) {
        autoSave.handleInput(id, description, updateStatusIndicator);
      }

      // log a new entry to clarity
      window.clarity("event", "captureAddress");
      postEventToSupabase(context === "fromShortcut" ? context : "");
    });
  });
};

function getDescription() {
  // return the description of the most recent card
  const capturedRanges = window.sharedState.capturedRanges;
  if (capturedRanges.length > 0) {
    return capturedRanges[capturedRanges.length - 1].description;
  }
  // if no card exists, default to placeholder
  return "Description of change...";
}
function getInitials() {
  return document.getElementById("initialsInput").value || "NU";
}
const updateCardContainer = async (focus_first = false) =>
  tryCatch(async () => {
    const cardContainer = document.getElementById("cardContainer");
    cardContainer.innerHTML = ""; // Clear existing cards

    if (window.sharedState.capturedRanges.length == 0) {
      cardContainer.innerHTML = `<p id="no-addresses-message">Click "Capture Address" to get started.</p>`;
      return;
    }

    window.sharedState.capturedRanges
      .slice()
      .reverse()
      .forEach((data, index) => {
        // Card root div
        const cardInstance = document.createElement("div");
        cardInstance.className = "card-instance";

        // Add insert button
        const sideBtn = document.createElement("button");
        sideBtn.id = `${data.id}-sidebtn`;
        setSideBtnState(data.id, getCardState(data.id), sideBtn, true);

        sideBtn.onclick = () => {
          if (window.sharedState.autoSave.enabled) return;
          insertSingleCard(window.sharedState.capturedRanges.length - 1 - index);
        };
        cardInstance.appendChild(sideBtn);

        // Card root div
        const card = document.createElement("div");
        card.className = "card";

        // Create card header
        const cardHeader = document.createElement("div");
        cardHeader.className = "card-header";

        // Create address wrapper
        const addressWrapper = document.createElement("div");
        addressWrapper.className = "card-address-wrapper";

        // Create sheet button
        const sheetButton = document.createElement("button");
        sheetButton.className = "card-address";
        sheetButton.textContent = data.sheet.length > 15 ? `${data.sheet.substring(0, 15)}...` : data.sheet;
        sheetButton.onclick = () => goToSheet(window.sharedState.capturedRanges.length - 1 - index);

        // Create range button
        const rangeButton = document.createElement("button");
        rangeButton.className = "card-address";
        rangeButton.textContent = data.address;
        rangeButton.onclick = () => goToAddress(window.sharedState.capturedRanges.length - 1 - index);

        // Add buttons to address wrapper
        addressWrapper.appendChild(sheetButton);
        addressWrapper.appendChild(rangeButton);

        // Create delete button
        const deleteButton = document.createElement("button");
        deleteButton.className = "card-delete-button";
        deleteButton.setAttribute("aria-label", "Delete");
        deleteButton.setAttribute("type", "button");
        deleteButton.innerHTML = `<img src="./assets/icon_delete_svg.svg" alt="Delete icon">`;
        deleteButton.onclick = () => deleteCard(window.sharedState.capturedRanges.length - 1 - index);

        // Add elements to card header
        cardHeader.appendChild(addressWrapper);
        cardHeader.appendChild(deleteButton);

        // Create card input
        const cardInput = document.createElement("div");
        cardInput.className = "card-input";
        cardInput.textContent = data.description;
        cardInput.setAttribute("contenteditable", "true");
        cardInput.setAttribute("placeholder", "Enter description...");
        cardInput.addEventListener("input", function () {
          window.sharedState.capturedRanges[window.sharedState.capturedRanges.length - 1 - index].description =
            this.textContent;
          window.sharedState.autoSave.handleInput(data.id, this.textContent, updateStatusIndicator);
        });
        cardInput.addEventListener("keydown", function (e) {
          if (e.key === "Enter" && !e.shiftKey) {
            this.blur();
            window.getSelection().removeAllRanges();

            // IMMEDIATELY EXECUTE UPON ENTER
            // Removing as no effect when debounce small
            // let autoSave = window.sharedState.autoSave;
            // if (autoSave.queue.has(data.id)) {
            //   const queuedData = autoSave.queue.get(data.id);

            //   // Execute the save immediately with the current value if there are pending changes
            //   autoSave.saveInput(data.id, queuedData.value, updateStatusIndicator);

            //   clearTimeout(queuedData.timeout); // Cancel any pending timeout for this input
            //   autoSave.queue.delete(data.id); // Remove from queue as we are saving now
            // } else {
            //   updateStatusIndicator(data.id, "saved"); // If no changes were pending, just mark as saved
            // }
          }
        });
        // focus on the first input if focus_first selected
        if (focus_first && index === 0) {
          setTimeout(() => {
            cardInput.focus();
            window.getSelection().selectAllChildren(cardInput);
          }, 0);
        }

        // Create card footer
        const cardFooter = document.createElement("div");
        cardFooter.className = "card-footer";
        cardFooter.innerHTML = `<p style="font-size:0.75rem">Use <span class="key-btn-text">shift+return</span> for a new line</p>`;
        cardFooter.style.display = "none"; // Initially hidden
        // Add event listeners to show/hide cardFooter based on cardInput focus
        cardInput.addEventListener("focus", function () {
          cardFooter.style.display = "block";
        });
        cardInput.addEventListener("blur", function () {
          cardFooter.style.display = "none";
        });

        // Add all elements to main card
        card.appendChild(cardHeader);
        card.appendChild(cardInput);
        card.appendChild(cardFooter);

        cardInstance.append(card);
        cardContainer.appendChild(cardInstance);
      });
  });

// Function to delete a card from the capturedRanges array and update the card container
const deleteCard = async (index) =>
  tryCatch(async () => {
    // Remove the card at the specified index from the capturedRanges array and assign it to a variable
    const data = window.sharedState.capturedRanges.splice(index, 1)[0];
    window.sharedState.autoSave.clearFromQueue(data.id);
    // Update the card container to reflect the changes in the capturedRanges array
    updateCardContainer();
  });

const insertSingleCard = async (index) =>
  tryCatch(async () => {
    if (window.sharedState.capturedRanges[index].state != "arrow") {
      // If not arrow (so either inserted or error), clicking button resets (instead of inserting again)
      window.sharedState.capturedRanges[index].state = "arrow";
      updateCardContainer();
    } else {
      insertAddress(index);
    }
  });

const insertAllCards = async () =>
  tryCatch(async () => {
    let i;
    for (i = 0; i < window.sharedState.capturedRanges.length; i++) {
      let [, error] = await insertAddress(i);
      if (error) break;
    }

    updateCardContainer();
  });

const deleteAllCards = async () =>
  tryCatch(async () => {
    window.sharedState.capturedRanges = [];
    window.sharedState.autoSave.clearQueue();
    updateCardContainer();
  });

const showTab = async (tabIndex) =>
  tryCatch(async () => {
    const tabs = ["home-tab", "address-clipper", "suggestions-tab"];

    // Hide all tabs
    tabs.forEach((id) => {
      document.getElementById(id).classList.remove("active");
    });

    // Show the selected tab
    document.getElementById(tabs[tabIndex - 1]).classList.add("active");
  });

const goToSheet = async (index) =>
  tryCatch(async () => {
    let data = window.sharedState.capturedRanges[index];

    await Excel.run(async (context) => {
      // Get the worksheet by name
      const sheet = context.workbook.worksheets.getItem(data.sheet);

      // Select the range in the Excel UI
      sheet.activate();

      // Synchronize the context to apply changes
      await context.sync();
    });
  });

const goToAddress = async (index) =>
  tryCatch(async () => {
    let data = window.sharedState.capturedRanges[index];

    await Excel.run(async (context) => {
      // Get the worksheet by name
      const sheet = context.workbook.worksheets.getItem(data.sheet);

      // Get the range using the specified address
      const range = sheet.getRange(data.address);

      // Select the range in the Excel UI
      range.select();

      // Synchronize the context to apply changes
      await context.sync();
    });
  });

function showPopup(message, isError = true) {
  const popup = document.createElement("div");
  popup.classList.add("popup");

  if (isError) {
    popup.classList.add("error-popup");
  }

  const messageElement = document.createElement("span");
  messageElement.textContent = message;

  const closeButton = document.createElement("button");
  closeButton.innerHTML = "&times;"; // Unicode character for "X"
  closeButton.addEventListener("click", () => {
    popup.classList.add("fade-out");
    setTimeout(() => {
      popup.remove();
    }, 500);
  });

  popup.appendChild(messageElement);
  popup.appendChild(closeButton);

  const container = document.querySelector(".popup-container");
  container.prepend(popup);

  setTimeout(() => {
    popup.classList.add("fade-out");
    setTimeout(() => {
      popup.remove();
    }, 500);
  }, 5000);
}

const insertAddress = async (index) =>
  tryCatch(async () => {
    await Excel.run(async (context) => {
      const activeCell = context.workbook.getActiveCell();

      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name"); // Explicitly load the 'name' property
      await context.sync(); // Sync to load the properties

      // Execute only if the active sheet is "Change log"
      if (activeSheet.name.toLowerCase() !== "change log") {
        throw Error("Unable to insert address. Destination sheet must be 'Change log'.");
      }

      const data = window.sharedState.capturedRanges[index];
      const initials = getInitials();

      // Create a horizontal array with initials and address
      const valueArray = constructEntryCells(data, initials);

      // insert address
      await insertDataAtCell(context, activeCell, valueArray);

      // Success!
      const nextRowCell = activeCell.getOffsetRange(1, 0);
      nextRowCell.select();

      await context.sync();

      setCardState(data.id, "saved");

      updateCardContainer();
    });
  });

function constructEntryCells(data, initials) {
  return [
    [
      data.id,
      data.sheet,
      getAddressLinkDynamic(data.sheet, data.address, data.fullAddress),
      data.description,
      initials,
    ],
  ];
}

function getAddressLinkDynamic(sht, addr, fullAddr) {
  // Construct the LET formula with HYPERLINK
  const linkFormula = `= LET(rng, ${fullAddr}, sht, TEXTAFTER(CELL("filename", rng), "]"), addr, IF(ROWS(rng) + COLUMNS(rng)=2, ADDRESS(ROW(rng), COLUMN(rng)), ADDRESS(MIN(ROW(rng)), MIN(COLUMN(rng))) & ":" & ADDRESS(MAX(ROW(rng)), MAX(COLUMN(rng)))), dynamic_link, HYPERLINK("#'" & sht & "'!" & addr, "↗️" & SUBSTITUTE(addr, "$", "")), IFERROR(dynamic_link, HYPERLINK("#'${sht}'!${addr}","[static!] ↗️${addr.replace(
    "$",
    ""
  )}")))`;

  return linkFormula;
}

// Configure the keyboard shortcut to open the task pane.
Office.actions.associate("ShowTaskpane", () => {
  return Office.addin
    .showAsTaskpane()
    .then(async () => {
      showTab(2);
      captureAddress("fromShortcut");
      return;
    })
    .catch((error) => {
      return error.code;
    });
});

// Configure the keyboard shortcut to close the task pane.
Office.actions.associate("HideTaskpane", () => {
  Office.addin
    .hide()
    .then(() => {
      return;
    })
    .catch((error) => {
      return error.code;
    });
});

async function getChangeLogSheet(context) {
  // get all worksheet names
  let sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  let changeLogSheet = sheets.items.find((sheet) => isChangeLog(sheet.name));

  if (!changeLogSheet) {
    // make sure autosave is turned off if on
    if (window.sharedState.autoSave.enabled) {
      const autosaveToggle = document.getElementById("autosaveToggle");
      autosaveToggle.checked = false;
      window.sharedState.autoSave.clearQueue();
      await handleAutosaveToggleChange(autosaveToggle, true);
      const error = new Error("Change log sheet not found. Disabling auto-save.");
      error.type = "ChangeLogNotFound";
      throw error;
    }
    throw new Error("Change log sheet not found.");
  }

  return changeLogSheet;
}

async function changeLogExists(context) {
  // get all worksheet names
  let sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  return sheets.items.find((sheet) => isChangeLog(sheet.name)) !== undefined;
}

function isChangeLog(sheetName) {
  return sheetName.toLowerCase().replace(/\s/g, "") === "changelog";
}

// Example usage:
const pushEntryToChangeLog = async (entryId, initials, data, pushIsUpdate = false) =>
  tryCatch(async () => {
    await Excel.run(async (context) => {
      // Find change log sheet

      let changeLogSheet = await getChangeLogSheet(context);

      // find used range
      let usedRange = changeLogSheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();

      // if blank insert template
      let freshLog = usedRange.isNullObject;
      if (freshLog) {
        // also set usedRange to newly inserted range
        usedRange = await insertTemplate(context, changeLogSheet);
      }

      // try to edit first (only if there were already entries)
      if (!freshLog) {
        // search for existing entry
        const idCol = usedRange.getColumn(0);
        idCol.load("values");
        await context.sync();

        const values = idCol.values;
        let foundCell;
        for (let i = 0; i < values.length; i++) {
          if (values[i][0] === entryId) {
            // Return the found cell as a range object
            foundCell = usedRange.getCell(i, 0); // Get the corresponding cell
          }
        }

        if (foundCell) {
          // FOUND CELL TO EDIT!
          // overwrite with new data (will skip null entries)
          await insertDataAtCell(context, foundCell, data, true);
          return;
        }
      }

      // INSERT
      // if you can't edit, insert new

      if (pushIsUpdate) {
        // caller asked for update, so data is incomplete
        // however, no cell was found to make update
        // therefore, find full data before continuing
        data = getCardEntryCells(entryId);
      }

      let insertBelowIndex;
      if (freshLog) {
        // if inserted into fresh template, default to first row
        insertBelowIndex = 1;
      } else {
        // otherwise, find best place to insert
        usedRange.load("rowIndex, rowCount, values");
        await context.sync();

        // set insertrow index to last
        insertBelowIndex = usedRange.rowCount - 1;

        let foundAny = false;

        for (let i = usedRange.values.length - 1; i >= 1; i--) {
          let rvals = usedRange.values[i];

          // reduce the default insertion point until you find anything (to make sure )
          if (!foundAny) {
            if (isEmpty([rvals])) {
              insertBelowIndex = i - 1; // its not this row so maybe its the next one, hence the '- 1'
            } else {
              foundAny = true;
            }
          }

          let val = usedRange.values[i][4];
          if (val === initials) {
            insertBelowIndex = i;
            break;
          }
        }
      }

      let rowBelow = usedRange.getRow(insertBelowIndex).getOffsetRange(1, 0);
      rowBelow.insert(Excel.InsertShiftDirection.down);

      let newRow = rowBelow.getOffsetRange(-1, 0);
      newRow.copyFrom(rowBelow, Excel.RangeCopyType.formats);

      // insert data at range
      let insertedRange = await insertDataAtCell(context, newRow, data);

      // assign colour to range
      let fillRange = insertedRange.getResizedRange(0, -1).getOffsetRange(0, 1);
      let [result, error] = await getFromLocalStorage("pwrups_fill_col");
      if (!error) fillRange.format.fill.color = result;

      let idCell = insertedRange.getCell(0);
      idCell.format.font.italic = true;
      idCell.format.font.size = 8;
      idCell.format.font.color = "#c2c2c2"; // Grey color in hexadecimal

      await context.sync();
    });
  });

async function insertDataAtCell(context, anchorRange, data, overwrite = false) {
  let rows = data.length;
  let cols = Math.max(...data.map((row) => row.length));

  let targetRange = anchorRange.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
  targetRange.load("formulas");
  await context.sync();

  if (overwrite || isEmpty(targetRange.formulas)) {
    // expand data to full rect (in case rows have different dimensions)
    const expandedData = data.map((row) => row.concat(Array(cols - row.length).fill(null)));

    const containsNull = expandedData.some((row) => row.includes(null));
    if (containsNull) {
      // Update values cell by cell, skipping over any cells where the data is null
      for (let i = 0; i < rows; i++) {
        // Iterate over each column in the current row
        for (let j = 0; j < cols; j++) {
          // Check if the current cell's data is not null
          if (expandedData[i][j] !== null) {
            // Update the cell's value
            targetRange.getCell(i, j).values = expandedData[i][j];
          }
        }
      }
    } else {
      // no null values to skip, so assign directly
      targetRange.values = expandedData;
    }

    await context.sync();

    console.log("Data inserted successfully.");
    return targetRange;
  } else {
    targetRange.select();
    console.log("Operation cancelled: Target range is not empty.");
    throw Error("Unable to insert address. Destination must be empty.");
  }
}

async function insertTemplate(context, sheet) {
  const templateData = [
    ["", "Change Log - Generated by FIECON Change Logger Powerup"],
    [
      "ID",
      "Sheet",
      "Cells",
      "Description",
      "Initials",
      "QC initials",
      "QC comment",
      "QC initials",
      "QC comment",
      "QC initials",
      "QC comment",
    ],
  ];

  let templateRange = await insertDataAtCell(context, sheet.getRange("A1"), templateData);

  let templateTitle = templateRange.getRow(0);
  templateTitle.format.font.italic = true;
  templateTitle.format.font.name = "Verdana";

  let templateContent = templateRange.getRow(1);
  templateContent.format.font.bold = true;
  templateContent.getResizedRange(0, -1).getOffsetRange(0, 1).format.fill.color = "#c4e6ff";
  templateContent.getResizedRange(0, -5).getOffsetRange(0, 5).format.fill.color = "#FBE2D5";

  templateContent.load("columnCount");
  await context.sync();

  // 3. Extend the number of rows in the range by 20 and set borders
  let extendedRange = templateContent.getResizedRange(20, 0);

  setBordersToBlack(extendedRange.getResizedRange(0, -1).getOffsetRange(0, 1));
  extendedRange.format.font.name = "Verdana";

  const caars = 5.725; // caars = completely arbitrary and random scale
  const columnWidths = [9, 18, 18, 45, 12, 12, 32, 12, 32, 12, 32];

  for (let i = 0; i < templateContent.columnCount; i++) {
    templateContent.getColumn(i).format.columnWidth = columnWidths[i] * caars;
  }

  await context.sync();

  // return entire range
  return templateTitle.getResizedRange(21, 0);
}

function setBordersToBlack(range) {
  let borders = range.format.borders;

  // Set all border sides to black
  borders.getItem("EdgeTop").color = "#000000";
  borders.getItem("EdgeBottom").color = "#000000";
  borders.getItem("EdgeLeft").color = "#000000";
  borders.getItem("EdgeRight").color = "#000000";
  borders.getItem("InsideVertical").color = "#000000";
  borders.getItem("InsideHorizontal").color = "#000000";

  // Set border style to continuous
  borders.getItem("EdgeTop").style = "Continuous";
  borders.getItem("EdgeBottom").style = "Continuous";
  borders.getItem("EdgeLeft").style = "Continuous";
  borders.getItem("EdgeRight").style = "Continuous";
  borders.getItem("InsideVertical").style = "Continuous";
  borders.getItem("InsideHorizontal").style = "Continuous";
}

function isEmpty(rngValues) {
  // alterantive to look into: apply following to rng.valueTypes - row.every((cell) => cell === Excel.RangeValueType.empty || cell === "" || cell === null)
  return rngValues.every((row) => row.every((cell) => cell === "" || cell === null));
}

function generateID(initials, address = "") {
  const baseChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const base = baseChars.length;
  let timestamp = Math.round((Date.now() - new Date("2025-01-01T00:00:00Z").getTime()) / 100);
  let result = "";

  while (timestamp > 0) {
    result = baseChars[timestamp % base] + result;
    timestamp = Math.floor(timestamp / base);
  }

  const randomChar = (str) => str[Math.floor(Math.random() * str.length)];

  // Removed address to simplify id
  // const extractAlphanumeric = (str) => str.replace(/[^a-z0-9]/gi, "");
  // return initials + extractAlphanumeric(address) + result + randomChar(baseChars) + randomChar(baseChars);

  return initials + result + randomChar(baseChars) + randomChar(baseChars);
}

// Function to update the status indicator
const updateStatusIndicator = async (entryId, status) =>
  tryCatch(async () => {
    console.log(`Setting ID ${entryId} to ${status}`);
    setSideBtnState(entryId, status);
    // if (indicatorElement) {
    //   indicatorElement.textContent = `Status: ${status}`;
    //   indicatorElement.className = `status-indicator ${status}`; // Add class for styling if needed
    // }
  });

const handleAutosaveToggleChange = async (checkbox, ignoreErr = false) =>
  tryCatch(async () => {
    await Excel.run(async (context) => {
      if (checkbox.checked) {
        // Simulate validation logic
        const hasLog = await changeLogExists(context);
        const hasInitials = document.getElementById("initialsInput").value.length > 0;

        if (!hasLog) {
          // revert the checkbox
          checkbox.checked = false;
          throw new Error("Auto-save not enabled. No 'Change log' sheet exists.");
        }

        if (!hasInitials) {
          // revert the checkbox
          checkbox.checked = false;
          throw new Error("Auto-save not enabled. Initials cannot be blank.");
        }

        // success!
        setAutosaveEnabled(true);
        return;
      }
      setAutosaveEnabled(false);
    });
  }, !ignoreErr);

const setAutosaveEnabled = async (enabled) =>
  tryCatch(async () => {
    let autoSave = window.sharedState.autoSave;

    // exit if no change
    if (enabled === autoSave.enabled) return;
    // enable auto-save!
    autoSave.enabled = enabled;

    if (enabled) {
      // ENABLE
      // Immediately push all existing to change log
      showPopup(
        "🛈 Auto-save enabled, syncing to Change Log. Deleting cards from this pane will not delete change log entries.",
        false
      );
      window.sharedState.capturedRanges.forEach((range) => {
        autoSave.handleInput(range.id, range.description, updateStatusIndicator);
      });
    } else {
      // DISABLE
      autoSave.clearQueue();
      window.sharedState.capturedRanges.forEach((range) => {
        range.state = "arrow";
      });
      console.log(window.sharedState.capturedRanges);
      await updateCardContainer();
    }
  });

function setSideBtnState(entryId, newState, buttonObj, forceUpdate = false) {
  // do nothing if no state change
  if (!forceUpdate && getCardState(entryId) === newState) {
    return;
  }

  // set optional param buttonObj if not given
  if (buttonObj === undefined) buttonObj = document.getElementById(`${entryId}-sidebtn`);

  // otherwise set state
  buttonObj.className = "card-insert" + (window.sharedState.autoSave.enabled ? "" : " clickable");
  setCardState(entryId, newState);

  switch (newState) {
    case "arrow":
      buttonObj.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="black" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="15 18 9 12 15 6"></polyline>
      </svg>`;
      break;
    case "saved":
      buttonObj.innerHTML = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#4CAF50" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="12" cy="12" r="11"/>
          <path d="M8 12l3 3 5-5"/>
      </svg>`;
      // set to complete
      buttonObj.className += " complete";
      break;
    case "pending":
      buttonObj.innerHTML = '<div class="spinner"></div>';
      break;
    case "error":
      buttonObj.innerHTML = "Err";
      break;
    case "none":
      buttonObj.innerHTML = "";
      break;
    default:
      buttonObj.innerHTML = "Err";
      console.error("Error or invalid state");
  }
}

const sampleFillColour = async () =>
  tryCatch(async () => {
    showPopup("Sampling fill colour from the active cell.", false);

    await Excel.run(async (context) => {
      const range = context.workbook.getActiveCell();
      range.load("format/fill/color");
      await context.sync();

      const cellColor = range.format.fill.color || "#FFFFFF";
      await setInLocalStorage("pwrups_fill_col", cellColor);

      const colourPicker = document.getElementById("fill-colour-input");
      colourPicker.value = cellColor;
    });
  });

function getRandomVibrantColor() {
  const h = Math.floor(Math.random() * 360);
  const s = 100;
  const l = 50;
  return hslToHex(h, s, l);
}

function hslToHex(h, s, l) {
  l /= 100;
  const a = (s * Math.min(l, 1 - l)) / 100;
  const f = (n) => {
    const k = (n + h / 30) % 12;
    const color = l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
    return Math.round(255 * color)
      .toString(16)
      .padStart(2, "0");
  };
  return `#${f(0)}${f(8)}${f(4)}`;
}

const postEventToSupabase = async (context = "") =>
  tryCatch(async () => {
    const apikey =
      "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ0d3ptc2tldmp0bWxicnZrZnl6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3MzM1ODA3NTAsImV4cCI6MjA0OTE1Njc1MH0.cOYXBDIaN1Sr3ldSoRq7zLbbABvQWJlcLR8MZOsLOS8";
    const url = "https://rtwzmskevjtmlbrvkfyz.supabase.co/rest/v1/powerup_actions";

    const headers = {
      apikey,
      Authorization: "Bearer " + apikey,
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    };

    const n_ranges = window.sharedState.capturedRanges.length;

    const data = {
      context,
      user_identifier: getInitials(),
      autosave_enabled: window.sharedState.autoSave.enabled,
      no_entries: n_ranges,
      // Calculate the average length of the descriptions of all items in capturedRanges, excluding the last one
      // divide by the number of elements (excluding the last one)
      avg_desc_len:
        window.sharedState.capturedRanges.slice(0, -1).reduce((acc, curr) => acc + curr.description.length, 0) / // sum up the lengths of all descriptions
        (n_ranges - 1),
    };

    const response = await fetch(url, {
      method: "POST",
      headers: headers,
      body: JSON.stringify(data),
    });

    if (!response.ok) throw new Error("Failed to log action");
  });
