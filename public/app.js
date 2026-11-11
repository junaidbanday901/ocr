/**
 * Voice Form & Excel Batch Generator
 * Front-end Logic, Web Speech API Controller, In-Memory Staging, and Excel Exporter
 */

// --- Application State ---
const fields = ['firstName', 'lastName', 'age', 'sex', 'phoneNo', 'address'];
const fieldLabels = {
  firstName: 'First Name',
  lastName: 'Last Name',
  age: 'Age',
  sex: 'Sex',
  phoneNo: 'Phone No',
  address: 'Address'
};

let currentMode = 'lineList'; // 'lineList' | 'nikshay'
let currentFieldIndex = 0;
let stagedRecords = [];
let stagedNikshayIds = [];
let isListening = false;
let recognition = null;
let restartTimeout = null;

// --- DOM Elements (Common & Header) ---
const homeLogoBtn = document.getElementById('homeLogoBtn');
const homeNavBtn = document.getElementById('homeNavBtn');
const headerModeBadge = document.getElementById('headerModeBadge');
const headerSubtitle = document.getElementById('headerSubtitle');
const modeSelectModal = document.getElementById('modeSelectModal');
const closeModeModalBtn = document.getElementById('closeModeModalBtn');
const selectLineListCard = document.getElementById('selectLineListCard');
const selectNikshayCard = document.getElementById('selectNikshayCard');

const lineListView = document.getElementById('lineListView');
const nikshayView = document.getElementById('nikshayView');
const activeFieldBanner = document.getElementById('activeFieldBanner');
const bannerFieldName = document.getElementById('bannerFieldName');
const bannerStep = document.getElementById('bannerStep');
const liveTranscriptText = document.getElementById('liveTranscriptText');
const toggleMicBtn = document.getElementById('toggleMicBtn');
const micBtnText = document.getElementById('micBtnText');
const speechStatusBadge = document.getElementById('speechStatusBadge');
const audioWaves = document.getElementById('audioWaves');

// --- Line List Elements ---
const firstNameInput = document.getElementById('firstName');
const lastNameInput = document.getElementById('lastName');
const ageInput = document.getElementById('age');
const sexSelect = document.getElementById('sex');
const phoneNoInput = document.getElementById('phoneNo');
const addressInput = document.getElementById('address');

const prevFieldBtn = document.getElementById('prevFieldBtn');
const nextFieldBtn = document.getElementById('nextFieldBtn');
const clearFormBtn = document.getElementById('clearFormBtn');
const addRowBtn = document.getElementById('addRowBtn');
const finalSubmitBtn = document.getElementById('finalSubmitBtn');
const clearAllBtn = document.getElementById('clearAllBtn');
const openFolderBtn = document.getElementById('openFolderBtn');

const recordCounter = document.getElementById('recordCounter');
const recordsTbody = document.getElementById('recordsTbody');
const emptyStateRow = document.getElementById('emptyStateRow');

// --- Nikshay Elements ---
const nikshayIdInput = document.getElementById('nikshayIdInput');
const addNikshayBtn = document.getElementById('addNikshayBtn');
const deleteLastNikshayBtn = document.getElementById('deleteLastNikshayBtn');
const clearNikshayInputBtn = document.getElementById('clearNikshayInputBtn');
const clearAllNikshayBtn = document.getElementById('clearAllNikshayBtn');
const finalSubmitNikshayBtn = document.getElementById('finalSubmitNikshayBtn');
const nikshayCounter = document.getElementById('nikshayCounter');
const nikshayTbody = document.getElementById('nikshayTbody');
const emptyNikshayRow = document.getElementById('emptyNikshayRow');
const nikshaySpeechStatusBadge = document.getElementById('nikshaySpeechStatusBadge');

const helpCommandsBtn = document.getElementById('helpCommandsBtn');
const commandsModal = document.getElementById('commandsModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const toastContainer = document.getElementById('toastContainer');

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
  setActiveField('firstName');
  initSpeechRecognition();
  setupEventListeners();
  loadExistingRecords();
  loadExistingNikshayRecords();
});

// --- Active Field Visual Tracking ---
let isProgrammaticFocus = false;

function setActiveField(fieldName) {
  const index = fields.indexOf(fieldName);
  if (index === -1) return;

  currentFieldIndex = index;

  // 1. Remove active state from all form groups
  document.querySelectorAll('.form-group').forEach((grp) => {
    grp.classList.remove('active-group', 'flash-active');
  });

  // 2. Add active state to target form group
  const targetGroup = document.getElementById(`group-${fieldName}`);
  if (targetGroup) {
    targetGroup.classList.add('active-group', 'flash-active');
    setTimeout(() => targetGroup.classList.remove('flash-active'), 600);
    targetGroup.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // 3. Focus corresponding input element without triggering recursion
  const inputEl = document.getElementById(fieldName);
  if (inputEl && document.activeElement !== inputEl) {
    isProgrammaticFocus = true;
    inputEl.focus();
    isProgrammaticFocus = false;
  }

  // 4. Update the top prominent banner
  bannerFieldName.textContent = fieldLabels[fieldName] || fieldName;
  bannerStep.textContent = `(Field ${currentFieldIndex + 1} of ${fields.length})`;

  // Play subtle feedback
  pulseBanner();
}

function nextField() {
  const nextIdx = (currentFieldIndex + 1) % fields.length;
  setActiveField(fields[nextIdx]);
}

function prevField() {
  const prevIdx = currentFieldIndex === 0 ? fields.length - 1 : currentFieldIndex - 1;
  setActiveField(fields[prevIdx]);
}

function pulseBanner() {
  const banner = document.getElementById('activeFieldBanner');
  banner.style.boxShadow = '0 0 30px rgba(56, 189, 248, 0.45)';
  setTimeout(() => {
    banner.style.boxShadow = '';
  }, 400);
}

// Quick selector for Sex
window.setSex = function (value) {
  sexSelect.value = value;
  showToast(`Sex set to ${value}`);
  nextField();
};

// --- Web Audio Confirmation Chime ---
function playChime(type = 'next') {
  try {
    const AudioContext = window.AudioContext || window.webkitAudioContext;
    if (!AudioContext) return;
    const ctx = new AudioContext();
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.connect(gain);
    gain.connect(ctx.destination);

    if (type === 'next') {
      osc.frequency.setValueAtTime(520, ctx.currentTime);
      osc.frequency.exponentialRampToValueAtTime(780, ctx.currentTime + 0.12);
      gain.gain.setValueAtTime(0.1, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.15);
      osc.start();
      osc.stop(ctx.currentTime + 0.15);
    } else if (type === 'add') {
      osc.frequency.setValueAtTime(440, ctx.currentTime);
      osc.frequency.setValueAtTime(660, ctx.currentTime + 0.08);
      osc.frequency.setValueAtTime(880, ctx.currentTime + 0.16);
      gain.gain.setValueAtTime(0.12, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.25);
      osc.start();
      osc.stop(ctx.currentTime + 0.25);
    }
  } catch (e) {
    // AudioContext blocked or not supported
  }
}

// Track recent commands to prevent duplicate execution from interim -> final
let lastExecutedCommand = '';
let lastExecutedTimestamp = 0;

function executeCommandOnce(commandKey, callback) {
  const now = Date.now();
  if (lastExecutedCommand === commandKey && now - lastExecutedTimestamp < 1200) {
    return false; // Already handled recently
  }
  lastExecutedCommand = commandKey;
  lastExecutedTimestamp = now;
  callback();
  return true;
}

// Clean speech string: remove punctuation, lowercase, collapse spaces
function cleanSpeechText(str) {
  if (!str) return '';
  return str
    .toLowerCase()
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?!"]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Check if a cleaned string is a "next" command
function isNextCommand(clean) {
  return /^(next|next field|go next|go to next|next one|forward|necks|neck|nex|skip)$/i.test(clean);
}

// Check if a cleaned string is a "back" command
function isBackCommand(clean) {
  return /^(back|previous|go back|prev)$/i.test(clean);
}

// Check if a cleaned string is an "add row" command
function isAddRowCommand(clean) {
  return /^(add row|next person|save row|save record|submit row|save|add)$/i.test(clean);
}

// --- Web Speech API Controller ---
function initSpeechRecognition() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

  if (!SpeechRecognition) {
    speechStatusBadge.textContent = 'Voice API Not Supported';
    speechStatusBadge.style.color = '#F43F5E';
    showToast('Web Speech API is not supported in this browser. Please use Google Chrome or Microsoft Edge.', 'error');
    return;
  }

  recognition = new SpeechRecognition();
  recognition.continuous = true;
  recognition.interimResults = true;
  recognition.lang = 'en-US';

  recognition.onstart = () => {
    isListening = true;
    updateMicUI(true);
  };

  recognition.onresult = (event) => {
    let interimTranscript = '';
    let finalTranscript = '';

    for (let i = event.resultIndex; i < event.results.length; ++i) {
      const transcriptPiece = event.results[i][0].transcript;
      if (event.results[i].isFinal) {
        finalTranscript += transcriptPiece;
      } else {
        interimTranscript += transcriptPiece;
      }
    }

    // --- Nikshay Mode Voice Routing ---
    if (currentMode === 'nikshay') {
      if (interimTranscript) {
        liveTranscriptText.textContent = `"${interimTranscript.trim()}"`;
        const interimClean = cleanSpeechText(interimTranscript);

        // Preview live extracted digits in input
        const previewDigits = extractNikshayDigits(interimTranscript);
        if (previewDigits) {
          nikshayIdInput.value = previewDigits;
        }

        // Fast-path commands for Nikshay mode
        if (/^(delete last|remove last|undo)$/i.test(interimClean)) {
          executeCommandOnce('nikshayDelete', () => {
            deleteLastNikshayId();
          });
          return;
        }

        if (/^(generate excel|final submit|save to excel|download excel)$/i.test(interimClean)) {
          executeCommandOnce('nikshayExcel', () => {
            finalSubmitNikshay();
          });
          return;
        }
      }

      if (finalTranscript) {
        liveTranscriptText.textContent = `"${finalTranscript.trim()}"`;
        processNikshayVoiceTranscript(finalTranscript.trim());
      }
      return;
    }

    // --- Line List Mode Voice Routing ---
    // Fast-path: Check interim transcript for standalone commands (like "next")
    if (interimTranscript) {
      const interimClean = cleanSpeechText(interimTranscript);
      liveTranscriptText.textContent = `"${interimTranscript.trim()}"`;

      if (isNextCommand(interimClean)) {
        executeCommandOnce('next', () => {
          nextField();
          playChime('next');
          showToast(`➡️ Advanced to ${fieldLabels[fields[currentFieldIndex]]}`, 'info');
        });
        return;
      }

      if (isBackCommand(interimClean)) {
        executeCommandOnce('back', () => {
          prevField();
          playChime('next');
          showToast(`⬅️ Returned to ${fieldLabels[fields[currentFieldIndex]]}`, 'info');
        });
        return;
      }

      if (isAddRowCommand(interimClean)) {
        executeCommandOnce('addRow', () => {
          addRowToMemory();
        });
        return;
      }
    }

    // Process final transcript in Line List mode
    if (finalTranscript) {
      liveTranscriptText.textContent = `"${finalTranscript.trim()}"`;
      processVoiceTranscript(finalTranscript.trim());
    }
  };

  recognition.onerror = (event) => {
    console.warn('Speech recognition error:', event.error);
    if (event.error === 'not-allowed') {
      speechStatusBadge.textContent = 'Mic Permission Denied';
      showToast('Microphone access denied. Please click the lock icon in the URL bar to allow microphone.', 'error');
      updateMicUI(false);
    }
  };

  recognition.onend = () => {
    // If intended to keep listening, auto-restart
    if (isListening) {
      clearTimeout(restartTimeout);
      restartTimeout = setTimeout(() => {
        try {
          recognition.start();
        } catch (e) {
          // Already active or error
        }
      }, 300);
    } else {
      updateMicUI(false);
    }
  };

  // Start listening by default
  startListening();
}

function startListening() {
  if (!recognition) return;
  try {
    isListening = true;
    recognition.start();
  } catch (err) {
    // Already running
  }
}

function stopListening() {
  if (!recognition) return;
  isListening = false;
  clearTimeout(restartTimeout);
  try {
    recognition.stop();
  } catch (err) {}
  updateMicUI(false);
}

function toggleListening() {
  if (isListening) {
    stopListening();
    showToast('Voice listening paused.');
  } else {
    startListening();
    showToast('Voice listening active. Speak freely!');
  }
}

function updateMicUI(active) {
  if (active) {
    toggleMicBtn.classList.add('active');
    micBtnText.textContent = 'Listening...';
    speechStatusBadge.textContent = 'Microphone Active';
    speechStatusBadge.style.color = '#34D399';
    if (nikshaySpeechStatusBadge) {
      nikshaySpeechStatusBadge.textContent = 'Microphone Active';
      nikshaySpeechStatusBadge.style.color = '#2DD4BF';
    }
    audioWaves.style.opacity = '1';
  } else {
    toggleMicBtn.classList.remove('active');
    micBtnText.textContent = 'Mic Paused';
    speechStatusBadge.textContent = 'Microphone Paused';
    speechStatusBadge.style.color = '#94A3B8';
    if (nikshaySpeechStatusBadge) {
      nikshaySpeechStatusBadge.textContent = 'Microphone Paused';
      nikshaySpeechStatusBadge.style.color = '#94A3B8';
    }
    audioWaves.style.opacity = '0.3';
  }
}

// --- Voice Transcript & Command Processor ---
function processVoiceTranscript(rawText) {
  const text = rawText.trim();
  const clean = cleanSpeechText(text);

  if (!clean) return;

  // 1. Standalone "Next" Command (e.g. "next", "next.", "next field", "forward")
  if (isNextCommand(clean)) {
    executeCommandOnce('next', () => {
      nextField();
      playChime('next');
      showToast(`➡️ Advanced to ${fieldLabels[fields[currentFieldIndex]]}`, 'info');
    });
    return;
  }

  // 2. Standalone "Back" Command
  if (isBackCommand(clean)) {
    executeCommandOnce('back', () => {
      prevField();
      playChime('next');
      showToast(`⬅️ Returned to ${fieldLabels[fields[currentFieldIndex]]}`, 'info');
    });
    return;
  }

  // 3. Standalone "Add Row" / "Next Person" Command
  if (isAddRowCommand(clean)) {
    executeCommandOnce('addRow', () => {
      addRowToMemory();
    });
    return;
  }

  // 4. Standalone "Clear" Command
  if (/^(clear|clear form|reset)$/i.test(clean)) {
    clearForm();
    showToast('Current form cleared');
    return;
  }

  // 5. Standalone "Generate Excel" Command
  if (/^(generate excel|final submit|save to excel|save excel|download excel)$/i.test(clean) || clean.includes('generate excel')) {
    finalSubmit();
    return;
  }

  // 6. Compound Command Ending with "Next" (e.g., "John next", "Smith next", "28 next", "Male next")
  // Allows user to speak value AND advance seamlessly in one breath!
  const compoundNextMatch = clean.match(/^(.*?)\s+\b(next|next field|go next|forward|necks|neck)\b$/i);
  if (compoundNextMatch && compoundNextMatch[1].trim()) {
    const valueToFill = compoundNextMatch[1].trim();
    fillCurrentField(valueToFill);
    executeCommandOnce('next', () => {
      nextField();
      playChime('next');
      showToast(`➡️ Filled & advanced to ${fieldLabels[fields[currentFieldIndex]]}`, 'info');
    });
    return;
  }

  // 7. Compound Command Ending with "Add Row" (e.g., "742 Evergreen Terrace add row")
  const compoundAddMatch = clean.match(/^(.*?)\s+\b(add row|next person|save row|save|add)\b$/i);
  if (compoundAddMatch && compoundAddMatch[1].trim()) {
    const valueToFill = compoundAddMatch[1].trim();
    fillCurrentField(valueToFill);
    executeCommandOnce('addRow', () => {
      addRowToMemory();
    });
    return;
  }

  // 8. Direct Field Jump Commands (e.g. "First Name John", "Age 28", "Phone 9876543210")
  if (clean.startsWith('first name ') || clean.startsWith('firstname ')) {
    const val = text.replace(/^first\s*name\s*/i, '').replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?!]/g, '').trim();
    setActiveField('firstName');
    fillCurrentField(val);
    return;
  }

  if (clean.startsWith('last name ') || clean.startsWith('lastname ')) {
    const val = text.replace(/^last\s*name\s*/i, '').replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?!]/g, '').trim();
    setActiveField('lastName');
    fillCurrentField(val);
    return;
  }

  if (clean.startsWith('age ')) {
    const val = text.replace(/^age\s*/i, '').trim();
    setActiveField('age');
    fillCurrentField(val);
    return;
  }

  if (clean.startsWith('sex ') || clean.startsWith('gender ')) {
    const val = text.replace(/^(sex|gender)\s*/i, '').trim();
    setActiveField('sex');
    fillCurrentField(val);
    return;
  }

  if (clean.startsWith('phone ') || clean.startsWith('phone number ') || clean.startsWith('mobile ')) {
    const val = text.replace(/^(phone\s*number|phone|mobile)\s*/i, '').trim();
    setActiveField('phoneNo');
    fillCurrentField(val);
    return;
  }

  if (clean.startsWith('address ')) {
    const val = text.replace(/^address\s*/i, '').trim();
    setActiveField('address');
    fillCurrentField(val);
    return;
  }

  // 9. Guard: NEVER fill a raw command word into the field
  if (isNextCommand(clean) || isBackCommand(clean) || isAddRowCommand(clean)) {
    return;
  }

  // 10. Otherwise, fill the currently active field with smart parsing
  fillCurrentField(text);
}

// --- Smart Field Parsers & Auto-Fill ---
function fillCurrentField(spokenText) {
  // Strip trailing punctuation
  const cleanSpoken = spokenText.replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?!"]+$/g, '').trim();
  if (!cleanSpoken) return;

  // Safety check: Don't fill command words
  const checkClean = cleanSpeechText(cleanSpoken);
  if (isNextCommand(checkClean) || isBackCommand(checkClean) || isAddRowCommand(checkClean)) {
    return;
  }

  const currentField = fields[currentFieldIndex];

  switch (currentField) {
    case 'firstName':
    case 'lastName':
      const formattedName = toTitleCase(cleanSpoken);
      document.getElementById(currentField).value = formattedName;
      showToast(`${fieldLabels[currentField]}: ${formattedName}`);
      break;

    case 'age':
      const parsedAge = parseAge(cleanSpoken);
      if (parsedAge) {
        ageInput.value = parsedAge;
        showToast(`Age: ${parsedAge}`);
      } else {
        ageInput.value = cleanSpoken;
      }
      break;

    case 'sex':
      const normalizedSex = parseSex(cleanSpoken);
      if (normalizedSex) {
        sexSelect.value = normalizedSex;
        showToast(`Sex: ${normalizedSex}`);
      }
      break;

    case 'phoneNo':
      const parsedPhone = parsePhoneNumber(cleanSpoken);
      phoneNoInput.value = parsedPhone;
      showToast(`Phone No: ${parsedPhone}`);
      break;

    case 'address':
      addressInput.value = toTitleCase(cleanSpoken);
      showToast(`Address updated`);
      break;
  }
}

// Utility: Title Case
function toTitleCase(str) {
  return str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}

// Utility: Parse Age from Words or Digits
function parseAge(text) {
  // Direct digits
  const digitMatch = text.match(/\d+/);
  if (digitMatch) {
    return parseInt(digitMatch[0], 10);
  }

  const wordMap = {
    one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
    eleven: 11, twelve: 12, thirteen: 13, fourteen: 14, fifteen: 15, sixteen: 16, seventeen: 17,
    eighteen: 18, nineteen: 19, twenty: 20, thirty: 30, forty: 40, fifty: 50, sixty: 60,
    seventy: 70, eighty: 80, ninety: 90
  };

  const words = text.toLowerCase().replace(/[^a-z\s]/g, '').split(/\s+/);
  let total = 0;
  for (let w of words) {
    if (wordMap[w] !== undefined) {
      total += wordMap[w];
    }
  }

  return total > 0 ? total : null;
}

// Utility: Parse Sex / Gender
function parseSex(text) {
  const lower = text.toLowerCase();
  if (lower.includes('female') || lower.includes('woman') || lower.includes('girl')) {
    return 'Female';
  }
  if (lower.includes('male') || lower.includes('man') || lower.includes('boy')) {
    return 'Male';
  }
  if (lower.includes('other') || lower.includes('non binary')) {
    return 'Other';
  }
  return null;
}

// Utility: Parse Phone Number
function parsePhoneNumber(text) {
  let lower = text.toLowerCase();

  // Replace common speech artifacts like "double five" or "triple nine"
  const digits = { zero: '0', oh: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9' };
  
  for (const [w, d] of Object.entries(digits)) {
    lower = lower.replace(new RegExp(`double\\s+${w}`, 'g'), `${d}${d}`);
    lower = lower.replace(new RegExp(`triple\\s+${w}`, 'g'), `${d}${d}${d}`);
    lower = lower.replace(new RegExp(`\\b${w}\\b`, 'g'), d);
  }

  // Extract all digits
  const extracted = lower.replace(/\D/g, '');
  return extracted || text;
}

// --- In-Memory Staging Queue ---
function addRowToMemory() {
  const firstName = firstNameInput.value.trim();
  const lastName = lastNameInput.value.trim();
  const age = ageInput.value.trim();
  const sex = sexSelect.value.trim();
  const phoneNo = phoneNoInput.value.trim();
  const address = addressInput.value.trim();

  // Validate at least one name or phone exists
  if (!firstName && !lastName && !phoneNo) {
    showToast('Please speak or enter at least a First Name or Last Name before adding row.', 'error');
    setActiveField('firstName');
    return;
  }

  const record = {
    id: stagedRecords.length + 1,
    timestamp: new Date().toLocaleString(),
    firstName,
    lastName,
    age,
    sex,
    phoneNo,
    address
  };

  stagedRecords.push(record);

  // Clear inputs for next person
  clearForm();

  // Reset focus back to First Name for smooth continuous entry!
  setActiveField('firstName');

  // Update Table UI & Counters
  renderTable();

  showToast(`Row ${stagedRecords.length} added to memory for ${firstName || 'Person'}! Ready for next person.`, 'success');
}

function clearForm() {
  firstNameInput.value = '';
  lastNameInput.value = '';
  ageInput.value = '';
  sexSelect.value = '';
  phoneNoInput.value = '';
  addressInput.value = '';
}

// --- Live Table Rendering ---
function renderTable() {
  // Update badge counter
  const count = stagedRecords.length;
  recordCounter.textContent = `${count} ${count === 1 ? 'Record' : 'Records'} Queued`;

  // Enable / disable final submission button
  if (count > 0) {
    finalSubmitBtn.removeAttribute('disabled');
    clearAllBtn.style.display = 'inline-block';
  } else {
    finalSubmitBtn.setAttribute('disabled', 'true');
    clearAllBtn.style.display = 'none';
  }

  // Clear table rows except empty state
  recordsTbody.innerHTML = '';

  if (count === 0) {
    recordsTbody.appendChild(emptyStateRow);
    return;
  }

  stagedRecords.forEach((record, index) => {
    const tr = document.createElement('tr');

    const sexBadgeClass = record.sex ? record.sex.toLowerCase() : '';
    const sexBadgeHtml = record.sex ? `<span class="td-badge-sex ${sexBadgeClass}">${record.sex}</span>` : '-';

    tr.innerHTML = `
      <td class="td-id">${index + 1}</td>
      <td class="td-name">${escapeHtml(record.firstName || '-')}</td>
      <td class="td-name">${escapeHtml(record.lastName || '-')}</td>
      <td>${record.age || '-'}</td>
      <td>${sexBadgeHtml}</td>
      <td class="td-phone">${escapeHtml(record.phoneNo || '-')}</td>
      <td style="max-width: 200px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${escapeHtml(record.address || '')}">${escapeHtml(record.address || '-')}</td>
      <td>
        <div class="td-actions">
          <button type="button" class="btn-row-action" title="Edit row" onclick="editRow(${index})">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
          </button>
          <button type="button" class="btn-row-action delete" title="Delete row" onclick="deleteRow(${index})">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
          </button>
        </div>
      </td>
    `;
    recordsTbody.appendChild(tr);
  });
}

window.deleteRow = function (index) {
  stagedRecords.splice(index, 1);
  renderTable();
  showToast('Row deleted from memory queue');
};

window.editRow = function (index) {
  const item = stagedRecords[index];
  if (!item) return;

  firstNameInput.value = item.firstName || '';
  lastNameInput.value = item.lastName || '';
  ageInput.value = item.age || '';
  sexSelect.value = item.sex || '';
  phoneNoInput.value = item.phoneNo || '';
  addressInput.value = item.address || '';

  // Remove from queue so when user clicks Add Row it re-adds updated version
  stagedRecords.splice(index, 1);
  renderTable();

  setActiveField('firstName');
  showToast('Row loaded into form for editing');
};

function clearAllStaged() {
  if (confirm('Are you sure you want to clear all in-memory records?')) {
    stagedRecords = [];
    renderTable();
    showToast('All staged records cleared');
  }
}

// --- Final Submission: Generate Excel (.xlsx) ---
async function finalSubmit() {
  if (stagedRecords.length === 0) {
    showToast('No records in memory to export. Please add at least one row.', 'error');
    return;
  }

  finalSubmitBtn.disabled = true;
  finalSubmitBtn.querySelector('span').textContent = 'Generating Excel...';

  try {
    // 1. Send batch to Node.js backend to save 'records.xlsx' directly to disk
    const response = await fetch('/api/generate-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows: stagedRecords })
    });

    const result = await response.json();

    if (result.success) {
      showToast(`Success! Generated records.xlsx with ${stagedRecords.length} records!`, 'success');
    } else {
      console.warn('Backend Excel generation notice:', result.message);
    }
  } catch (err) {
    console.warn('Backend fetch failed, using client-side SheetJS fallback:', err);
  }

  // 2. Also trigger instant in-browser download via SheetJS (xlsx) for immediate user access
  try {
    downloadExcelInBrowser(stagedRecords);
  } catch (clientErr) {
    console.error('Client Excel download error:', clientErr);
  }

  finalSubmitBtn.disabled = false;
  finalSubmitBtn.querySelector('span').textContent = 'Generate & Save Excel';
}

// Client-side SheetJS fallback/dual download
function downloadExcelInBrowser(data) {
  if (typeof XLSX === 'undefined') return;

  const exportData = data.map((r, i) => ({
    '#': i + 1,
    'Entry Time': r.timestamp,
    'First Name': r.firstName,
    'Last Name': r.lastName,
    'Age': r.age,
    'Sex': r.sex,
    'Phone No': r.phoneNo,
    'Address': r.address
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Voice Records');
  XLSX.writeFile(workbook, 'records.xlsx');
}

// Open folder via server endpoint
async function openWorkspaceFolder() {
  try {
    const res = await fetch('/api/open-folder', { method: 'POST' });
    const data = await res.json();
    if (data.success) {
      showToast('Opened project folder in File Explorer');
    }
  } catch (err) {
    showToast('Workspace folder is located at your desktop ocr directory');
  }
}

// Load existing records from server if records.xlsx already exists
async function loadExistingRecords() {
  try {
    const res = await fetch('/api/records');
    const data = await res.json();
    if (data.success && Array.isArray(data.rows) && data.rows.length > 0) {
      stagedRecords = data.rows;
      renderTable();
      showToast(`Loaded ${data.rows.length} existing records from records.xlsx`);
    }
  } catch (err) {
    // No previous file, fine
  }
}

// --- Toast Notification Helper ---
function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;

  let icon = '📢';
  if (type === 'success') icon = '✅';
  if (type === 'error') icon = '⚠️';

  toast.innerHTML = `<span>${icon}</span> <span>${escapeHtml(message)}</span>`;
  toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3500);
}

// Helper: Escape HTML
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// --- Mode Selection & Navigation ---
function openModeModal() {
  modeSelectModal.classList.add('open');
}

function closeModeModal() {
  modeSelectModal.classList.remove('open');
}

function switchMode(mode) {
  currentMode = mode;
  closeModeModal();

  if (mode === 'lineList') {
    lineListView.style.display = 'grid';
    activeFieldBanner.style.display = 'flex';
    nikshayView.style.display = 'none';

    headerModeBadge.textContent = 'LINE LIST';
    headerModeBadge.className = 'badge-pro';
    headerSubtitle.innerHTML = 'Continuous Voice Entry &bull; In-Memory Staging &bull; Excel Generator';

    setActiveField('firstName');
    showToast('Switched to Line List Generator (Demographic Form)', 'info');
  } else if (mode === 'nikshay') {
    lineListView.style.display = 'none';
    activeFieldBanner.style.display = 'none';
    nikshayView.style.display = 'grid';

    headerModeBadge.textContent = 'NIKSHAY IDS';
    headerModeBadge.className = 'badge-pro teal-badge';
    headerSubtitle.innerHTML = 'Rapid Continuous Voice Entry &bull; Nikshay IDs &bull; Excel Generator';

    setTimeout(() => {
      nikshayIdInput.focus();
    }, 150);
    showToast('Switched to Nikshay IDs Fast Voice Entry Mode', 'info');
  }
}
window.switchMode = switchMode;

// --- Nikshay Spoken Number Parser & Processor ---
function extractNikshayDigits(text) {
  if (!text) return '';

  const clean = text.toLowerCase()
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()?!]/g, ' ')
    .trim();

  // If text is already pure digits
  const directDigits = clean.replace(/\s+/g, '');
  if (/^\d+$/.test(directDigits)) {
    return directDigits;
  }

  const wordsToDigits = {
    'zero': '0', 'oh': '0', 'o': '0', 'null': '0', 'nil': '0',
    'one': '1', 'won': '1',
    'two': '2', 'to': '2', 'too': '2',
    'three': '3',
    'four': '4', 'for': '4', 'fore': '4',
    'five': '5',
    'six': '6',
    'seven': '7',
    'eight': '8', 'ate': '8',
    'nine': '9'
  };

  const tensTeens = {
    'ten': '10', 'eleven': '11', 'twelve': '12', 'thirteen': '13', 'fourteen': '14',
    'fifteen': '15', 'sixteen': '16', 'seventeen': '17', 'eighteen': '18', 'nineteen': '19',
    'twenty': '20', 'thirty': '30', 'forty': '40', 'fifty': '50', 'sixty': '60',
    'seventy': '70', 'eighty': '80', 'ninety': '90', 'hundred': '00', 'thousand': '000'
  };

  const tokens = clean.split(/\s+/);
  let resultDigits = '';

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    // Handle multipliers like "double five" or "triple eight"
    if ((token === 'double' || token === 'trouble' || token === 'bubble') && i + 1 < tokens.length) {
      const next = tokens[i + 1];
      const digit = wordsToDigits[next] || (/^\d$/.test(next) ? next : null);
      if (digit) {
        resultDigits += digit + digit;
        i++;
        continue;
      }
    }

    if (token === 'triple' && i + 1 < tokens.length) {
      const next = tokens[i + 1];
      const digit = wordsToDigits[next] || (/^\d$/.test(next) ? next : null);
      if (digit) {
        resultDigits += digit + digit + digit;
        i++;
        continue;
      }
    }

    if (/^\d+$/.test(token)) {
      resultDigits += token;
    } else if (wordsToDigits[token] !== undefined) {
      resultDigits += wordsToDigits[token];
    } else if (tensTeens[token] !== undefined) {
      resultDigits += tensTeens[token];
    }
  }

  return resultDigits;
}

function processNikshayVoiceTranscript(rawText) {
  const text = rawText.trim();
  const clean = cleanSpeechText(text);
  if (!clean) return;

  // 1. Generate Excel Command
  if (/^(generate excel|final submit|save to excel|save excel|download excel|export excel)$/i.test(clean) || clean.includes('generate excel')) {
    executeCommandOnce('nikshayExcel', () => {
      finalSubmitNikshay();
    });
    return;
  }

  // 2. Delete Last Command
  if (/^(delete last|delete|remove last|remove|undo)$/i.test(clean)) {
    executeCommandOnce('nikshayDelete', () => {
      deleteLastNikshayId();
    });
    return;
  }

  // 3. Clear All Command
  if (/^(clear all|clear list|delete all)$/i.test(clean)) {
    clearAllNikshayIds();
    return;
  }

  // 4. Clear Input Field Command
  if (/^(clear|clear input|clear field|reset)$/i.test(clean)) {
    nikshayIdInput.value = '';
    showToast('Nikshay input cleared');
    return;
  }

  // 5. Number Extraction & Auto-Add to List
  const digits = extractNikshayDigits(text);
  if (digits && digits.length >= 2) {
    nikshayIdInput.value = digits;
    addNikshayId(digits);
  } else if (/^\d+$/.test(clean)) {
    nikshayIdInput.value = clean;
    addNikshayId(clean);
  }
}

// --- Nikshay Staging & Table Management ---
function addNikshayId(idString) {
  const cleanId = String(idString || '').trim();
  if (!cleanId) {
    showToast('Please speak or type a valid Nikshay ID number', 'error');
    return;
  }

  const isDuplicate = stagedNikshayIds.some((item) => item.nikshayId === cleanId);

  stagedNikshayIds.push({
    id: stagedNikshayIds.length + 1,
    timestamp: new Date().toLocaleString(),
    nikshayId: cleanId
  });

  renderNikshayTable();
  nikshayIdInput.value = '';
  nikshayIdInput.focus();
  playChime('add');

  if (isDuplicate) {
    showToast(`⚠️ Added Nikshay ID: ${cleanId} (Note: ID already in list)`, 'info');
  } else {
    showToast(`✅ Added Nikshay ID: ${cleanId}`, 'success');
  }
}

function deleteNikshayId(index) {
  if (index >= 0 && index < stagedNikshayIds.length) {
    const removed = stagedNikshayIds.splice(index, 1)[0];
    // Re-index remaining IDs
    stagedNikshayIds.forEach((item, idx) => {
      item.id = idx + 1;
    });
    renderNikshayTable();
    showToast(`🗑️ Removed Nikshay ID: ${removed.nikshayId}`, 'info');
  }
}
window.deleteNikshayId = deleteNikshayId;

function deleteLastNikshayId() {
  if (stagedNikshayIds.length === 0) {
    showToast('No Nikshay IDs to delete', 'info');
    return;
  }
  const removed = stagedNikshayIds.pop();
  renderNikshayTable();
  showToast(`🗑️ Removed latest Nikshay ID: ${removed.nikshayId}`, 'info');
}

function clearAllNikshayIds() {
  if (stagedNikshayIds.length === 0) {
    showToast('List is already empty');
    return;
  }
  if (confirm('Are you sure you want to clear all queued Nikshay IDs?')) {
    stagedNikshayIds = [];
    renderNikshayTable();
    showToast('Cleared all queued Nikshay IDs');
  }
}

function renderNikshayTable() {
  const count = stagedNikshayIds.length;
  nikshayCounter.textContent = `${count} ID${count === 1 ? '' : 's'} Queued`;

  // Remove existing dynamic rows
  const dynamicRows = nikshayTbody.querySelectorAll('tr:not(#emptyNikshayRow)');
  dynamicRows.forEach((row) => row.remove());

  if (count === 0) {
    emptyNikshayRow.style.display = '';
    clearAllNikshayBtn.style.display = 'none';
    finalSubmitNikshayBtn.disabled = true;
    return;
  }

  emptyNikshayRow.style.display = 'none';
  clearAllNikshayBtn.style.display = 'inline-block';
  finalSubmitNikshayBtn.disabled = false;

  stagedNikshayIds.forEach((item, index) => {
    const tr = document.createElement('tr');
    tr.className = 'staged-row';

    tr.innerHTML = `
      <td class="td-id">${index + 1}</td>
      <td style="font-size:0.85rem; color:var(--text-dim);">${escapeHtml(item.timestamp)}</td>
      <td class="td-nikshay">${escapeHtml(item.nikshayId)}</td>
      <td class="td-actions">
        <button type="button" class="btn-row-action delete" title="Delete ID" onclick="deleteNikshayId(${index})">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
        </button>
      </td>
    `;
    nikshayTbody.appendChild(tr);
  });
}

// Final Nikshay Excel Submission
async function finalSubmitNikshay() {
  if (stagedNikshayIds.length === 0) {
    showToast('No Nikshay IDs queued to generate Excel!', 'error');
    return;
  }

  finalSubmitNikshayBtn.disabled = true;
  finalSubmitNikshayBtn.querySelector('span').textContent = 'Generating Excel...';

  try {
    const response = await fetch('/api/generate-nikshay-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rows: stagedNikshayIds })
    });

    const data = await response.json();

    if (data.success) {
      playChime('submit');
      showToast(`🎉 ${data.message}`, 'success');
      window.location.href = '/api/download-nikshay';
    } else {
      throw new Error(data.message || 'Server returned failure');
    }
  } catch (err) {
    console.warn('Backend Nikshay fetch failed, using client-side SheetJS fallback:', err);
    try {
      downloadNikshayExcelInBrowser(stagedNikshayIds);
      playChime('submit');
      showToast(`🎉 Successfully generated nikshay_records.xlsx with ${stagedNikshayIds.length} IDs!`, 'success');
    } catch (clientErr) {
      console.error('Client Excel download error:', clientErr);
      showToast('Failed to generate Excel file: ' + err.message, 'error');
    }
  }

  finalSubmitNikshayBtn.disabled = false;
  finalSubmitNikshayBtn.querySelector('span').textContent = 'Generate & Save Excel';
}

function downloadNikshayExcelInBrowser(data) {
  if (typeof XLSX === 'undefined') return;

  const exportData = data.map((r, i) => ({
    '#': i + 1,
    'Entry Time': r.timestamp,
    'Nikshay ID': r.nikshayId
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Nikshay IDs');
  XLSX.writeFile(workbook, 'nikshay_records.xlsx');
}

async function loadExistingNikshayRecords() {
  try {
    const res = await fetch('/api/nikshay-records');
    const data = await res.json();
    if (data.success && Array.isArray(data.rows) && data.rows.length > 0) {
      stagedNikshayIds = data.rows;
      renderNikshayTable();
      showToast(`Loaded ${data.rows.length} existing records from nikshay_records.xlsx`);
    }
  } catch (err) {
    // No previous file, fine
  }
}

// --- Setup Event Listeners ---
function setupEventListeners() {
  toggleMicBtn.addEventListener('click', toggleListening);
  prevFieldBtn.addEventListener('click', prevField);
  nextFieldBtn.addEventListener('click', nextField);
  clearFormBtn.addEventListener('click', clearForm);
  addRowBtn.addEventListener('click', addRowToMemory);
  finalSubmitBtn.addEventListener('click', finalSubmit);
  clearAllBtn.addEventListener('click', clearAllStaged);
  openFolderBtn.addEventListener('click', openWorkspaceFolder);

  // Home & Mode selection handlers
  if (homeLogoBtn) homeLogoBtn.addEventListener('click', openModeModal);
  if (homeNavBtn) homeNavBtn.addEventListener('click', openModeModal);
  if (closeModeModalBtn) closeModeModalBtn.addEventListener('click', closeModeModal);

  if (selectLineListCard) {
    selectLineListCard.addEventListener('click', () => switchMode('lineList'));
    selectLineListCard.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') switchMode('lineList');
    });
  }

  if (selectNikshayCard) {
    selectNikshayCard.addEventListener('click', () => switchMode('nikshay'));
    selectNikshayCard.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') switchMode('nikshay');
    });
  }

  if (modeSelectModal) {
    modeSelectModal.addEventListener('click', (e) => {
      if (e.target === modeSelectModal) closeModeModal();
    });
  }

  // Nikshay form & control handlers
  if (addNikshayBtn) {
    addNikshayBtn.addEventListener('click', () => {
      addNikshayId(nikshayIdInput.value);
    });
  }

  if (nikshayIdInput) {
    nikshayIdInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        addNikshayId(nikshayIdInput.value);
      }
    });
  }

  if (deleteLastNikshayBtn) deleteLastNikshayBtn.addEventListener('click', deleteLastNikshayId);
  if (clearNikshayInputBtn) {
    clearNikshayInputBtn.addEventListener('click', () => {
      nikshayIdInput.value = '';
      nikshayIdInput.focus();
    });
  }
  if (clearAllNikshayBtn) clearAllNikshayBtn.addEventListener('click', clearAllNikshayIds);
  if (finalSubmitNikshayBtn) finalSubmitNikshayBtn.addEventListener('click', finalSubmitNikshay);

  // Keyboard navigation shortcuts
  document.addEventListener('keydown', (e) => {
    // Ctrl + Enter or Cmd + Enter
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      if (currentMode === 'lineList') {
        addRowToMemory();
      } else if (currentMode === 'nikshay') {
        addNikshayId(nikshayIdInput.value);
      }
    }
  });

  // Modal handlers
  helpCommandsBtn.addEventListener('click', () => {
    commandsModal.classList.add('open');
  });

  closeModalBtn.addEventListener('click', () => {
    commandsModal.classList.remove('open');
  });

  commandsModal.addEventListener('click', (e) => {
    if (e.target === commandsModal) {
      commandsModal.classList.remove('open');
    }
  });

  // Manual click on inputs updates active field
  fields.forEach((fieldName) => {
    const el = document.getElementById(fieldName);
    if (el) {
      el.addEventListener('focus', () => {
        if (!isProgrammaticFocus) {
          setActiveField(fieldName);
        }
      });
    }
  });
}
