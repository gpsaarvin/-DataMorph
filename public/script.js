/**
 * ============================================
 *  DataMorphp — Smart PDF to Excel Converter
 *  Client-side Script
 * ============================================
 *  Handles drag-and-drop, file selection,
 *  upload progress simulation, conversion
 *  request, and success/error feedback.
 * ============================================
 */

(function () {
  'use strict';

  // ── DOM References ─────────────────────────────
  const dropzone       = document.getElementById('dropzone');
  const fileInput      = document.getElementById('fileInput');
  const fileInfo       = document.getElementById('fileInfo');
  const fileName       = document.getElementById('fileName');
  const removeFileBtn  = document.getElementById('removeFile');
  const convertBtn     = document.getElementById('convertBtn');
  const btnSpinner     = document.getElementById('btnSpinner');
  const progressWrapper= document.getElementById('progressWrapper');
  const progressFill   = document.getElementById('progressFill');
  const progressText   = document.getElementById('progressText');
  const status         = document.getElementById('status');

  /** Currently selected PDF file */
  let selectedFile = null;

  // ── Drag & Drop Events ─────────────────────────
  const preventDefaults = (e) => {
    e.preventDefault();
    e.stopPropagation();
  };

  ['dragenter', 'dragover', 'dragleave', 'drop'].forEach((evt) => {
    dropzone.addEventListener(evt, preventDefaults);
  });

  ['dragenter', 'dragover'].forEach((evt) => {
    dropzone.addEventListener(evt, () => dropzone.classList.add('drag-over'));
  });

  ['dragleave', 'drop'].forEach((evt) => {
    dropzone.addEventListener(evt, () => dropzone.classList.remove('drag-over'));
  });

  dropzone.addEventListener('drop', (e) => {
    const files = e.dataTransfer.files;
    if (files.length > 0) handleFileSelect(files[0]);
  });

  // Click on the dropzone (outside the browse button) should also open file dialog
  dropzone.addEventListener('click', (e) => {
    if (e.target.tagName !== 'LABEL' && e.target.tagName !== 'INPUT') {
      fileInput.click();
    }
  });

  // ── Browse File ────────────────────────────────
  fileInput.addEventListener('change', () => {
    if (fileInput.files.length > 0) handleFileSelect(fileInput.files[0]);
  });

  // ── File Selection Handler ─────────────────────
  function handleFileSelect(file) {
    // Validate PDF
    if (file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
      showStatus('Please select a valid PDF file.', 'error');
      return;
    }

    // Validate size (25 MB)
    if (file.size > 25 * 1024 * 1024) {
      showStatus('File is too large. Maximum size is 25 MB.', 'error');
      return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.classList.add('show');
    convertBtn.disabled = false;
    clearStatus();
    hideProgress();
  }

  // ── Remove File ────────────────────────────────
  removeFileBtn.addEventListener('click', resetState);

  function resetState() {
    selectedFile = null;
    fileInput.value = '';
    fileName.textContent = 'No file selected';
    fileInfo.classList.remove('show');
    convertBtn.disabled = true;
    clearStatus();
    hideProgress();
    setButtonLoading(false);
  }

  // ── Convert Button ─────────────────────────────
  convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    try {
      setButtonLoading(true);
      clearStatus();
      showProgress();
      showStatus('Processing PDF... This may take up to 2 minutes for complex documents.', 'info');
      simulateProgress();

      // Build form data
      const formData = new FormData();
      formData.append('pdfFile', selectedFile);

      // Send to backend
      const response = await fetch('/upload', {
        method: 'POST',
        body: formData,
      });

      // Complete the progress bar
      completeProgress();

      if (!response.ok) {
        const err = await response.json().catch(() => ({ error: 'Conversion failed.' }));
        throw new Error(err.error || 'Something went wrong.');
      }

      // Read the blob and trigger download
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'StudentData.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      // Show success
      showStatus('Conversion successful! Your download has started.', 'success');
      showSuccessPopup();
    } catch (err) {
      completeProgress();
      showStatus(err.message || 'An unexpected error occurred.', 'error');
    } finally {
      setButtonLoading(false);
    }
  });

  // ── Progress Simulation ────────────────────────
  let progressInterval = null;

  function showProgress() {
    progressWrapper.classList.add('show');
    progressFill.style.width = '0%';
    progressText.textContent = '0%';
  }

  function hideProgress() {
    progressWrapper.classList.remove('show');
    progressFill.style.width = '0%';
    progressText.textContent = '0%';
    if (progressInterval) clearInterval(progressInterval);
  }

  function simulateProgress() {
    let value = 0;
    if (progressInterval) clearInterval(progressInterval);
    const steps = [
      { msg: 'Reading PDF...', at: 5 },
      { msg: 'Rendering pages for OCR...', at: 15 },
      { msg: 'Running text recognition...', at: 30 },
      { msg: 'Extracting student data...', at: 60 },
      { msg: 'Building Excel file...', at: 85 },
    ];
    let stepIdx = 0;
    progressInterval = setInterval(() => {
      // Slower progress for OCR-heavy processing
      const increment = Math.max(0.3, (90 - value) * 0.03);
      value = Math.min(value + increment, 90);
      updateProgress(value);
      if (stepIdx < steps.length && value >= steps[stepIdx].at) {
        showStatus(steps[stepIdx].msg, 'info');
        stepIdx++;
      }
    }, 200);
  }

  function completeProgress() {
    if (progressInterval) clearInterval(progressInterval);
    updateProgress(100);
  }

  function updateProgress(value) {
    const rounded = Math.round(value);
    progressFill.style.width = `${rounded}%`;
    progressText.textContent = `${rounded}%`;
  }

  // ── Button Loading State ───────────────────────
  function setButtonLoading(loading) {
    if (loading) {
      convertBtn.classList.add('btn--loading');
      convertBtn.disabled = true;
    } else {
      convertBtn.classList.remove('btn--loading');
      convertBtn.disabled = !selectedFile;
    }
  }

  // ── Status Messages ────────────────────────────
  function showStatus(message, type) {
    status.textContent = message;
    status.className = `status status--${type}`;
  }

  function clearStatus() {
    status.textContent = '';
    status.className = 'status';
  }

  // ── Success Popup ──────────────────────────────
  function showSuccessPopup() {
    // Remove any existing popup
    const existing = document.querySelector('.success-popup');
    if (existing) existing.remove();

    const popup = document.createElement('div');
    popup.className = 'success-popup';
    popup.innerHTML = `
      <div class="success-popup__card">
        <div class="success-popup__icon">
          <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24"
               fill="none" stroke="currentColor" stroke-width="2.5"
               stroke-linecap="round" stroke-linejoin="round">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
        </div>
        <h3 class="success-popup__title">Conversion Complete!</h3>
        <p class="success-popup__text">Your Excel file has been generated and the download has started automatically.</p>
        <button class="btn btn--primary" id="popupClose">Done</button>
      </div>
    `;

    document.body.appendChild(popup);

    // Close handlers
    const closePopup = () => {
      popup.style.animation = 'fadeIn 0.2s ease reverse';
      setTimeout(() => popup.remove(), 200);
      resetState();
    };

    popup.querySelector('#popupClose').addEventListener('click', closePopup);
    popup.addEventListener('click', (e) => {
      if (e.target === popup) closePopup();
    });
  }
})();
