(() => {
  const dropZone    = document.getElementById('dropZone');
  const fileInput   = document.getElementById('fileInput');
  const fileSelected = document.getElementById('fileSelected');
  const fileName    = document.getElementById('fileName');
  const fileSize    = document.getElementById('fileSize');
  const removeFile  = document.getElementById('removeFile');
  const convertBtn  = document.getElementById('convertBtn');
  const btnLabel    = convertBtn.querySelector('.btn-label');
  const btnSpinner  = convertBtn.querySelector('.btn-spinner');
  const ocrOption   = document.getElementById('ocrOption');
  const ocrToggle   = document.getElementById('ocrToggle');
  const progressWrap = document.getElementById('progressWrap');
  const progressBar  = document.getElementById('progressBar');
  const alertError   = document.getElementById('alertError');
  const alertSuccess = document.getElementById('alertSuccess');
  const errorMsg     = document.getElementById('errorMsg');
  const tabs         = document.querySelectorAll('.tab');

  let selectedFile = null;
  let currentFormat = 'docx';

  // ── TABS ──────────────────────────────────────────
  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentFormat = tab.dataset.format;
      ocrOption.style.display = currentFormat === 'docx' ? 'block' : 'none';
    });
  });

  // ── DROP ZONE ─────────────────────────────────────
  dropZone.addEventListener('click', () => fileInput.click());

  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file) setFile(file);
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) setFile(fileInput.files[0]);
  });

  removeFile.addEventListener('click', clearFile);

  // ── FILE HELPERS ──────────────────────────────────
  function setFile(file) {
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      showError('Somente arquivos PDF são aceitos.');
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      showError('Arquivo excede o limite de 50 MB.');
      return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatSize(file.size);
    dropZone.style.display = 'none';
    fileSelected.style.display = 'block';
    convertBtn.disabled = false;
    hideAlerts();
  }

  function clearFile() {
    selectedFile = null;
    fileInput.value = '';
    dropZone.style.display = 'block';
    fileSelected.style.display = 'none';
    convertBtn.disabled = true;
    hideAlerts();
    resetProgress();
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
  }

  // ── CONVERT ───────────────────────────────────────
  convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    hideAlerts();
    setLoading(true);

    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('format', currentFormat);
    if (currentFormat === 'docx') {
      formData.append('ocr', ocrToggle.checked ? 'true' : 'false');
    }

    startFakeProgress();

    try {
      const successText = alertSuccess.querySelector('span');
      if (successText) {
        successText.textContent = 'Arquivo convertido com sucesso! O download iniciará automaticamente.';
      }
      let res = await fetch('/convert', { method: 'POST', body: formData });

      if (!res.ok) {
        const contentType = (res.headers.get('Content-Type') || '').toLowerCase();
        let errMsg = '';

        if (contentType.includes('application/json')) {
          const data = await res.json().catch(() => ({}));
          errMsg = (data && data.error) ? String(data.error) : '';
        } else {
          const txt = await res.text().catch(() => '');
          if (txt) {
            // Remove tags HTML para exibir mensagem legível no alerta.
            errMsg = txt.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
          }
        }

        if (!errMsg) {
          errMsg = `Falha na conversão (HTTP ${res.status}${res.statusText ? ` - ${res.statusText}` : ''}).`;
        }
        throw new Error(errMsg);
      }

      finishProgress();

      const blob = await res.blob();
      const disposition = res.headers.get('Content-Disposition') || '';
      const nameMatch = disposition.match(/filename[^;=\n]*=['"]?([^'";\n]+)/);
      const dlName = nameMatch ? nameMatch[1] : `convertido.${currentFormat}`;

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = dlName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      alertSuccess.style.display = 'flex';
    } catch (err) {
      showError((err && err.message) ? err.message : 'Falha na conversão. Tente novamente.');
      resetProgress();
    } finally {
      setLoading(false);
    }
  });

  // ── PROGRESS FAKE ──────────────────────────────────
  let progressInterval = null;
  let progressVal = 0;

  function startFakeProgress() {
    progressVal = 0;
    progressWrap.style.display = 'block';
    progressBar.style.width = '0%';
    progressInterval = setInterval(() => {
      if (progressVal < 85) {
        progressVal += Math.random() * 6;
        progressBar.style.width = Math.min(progressVal, 85) + '%';
      }
    }, 400);
  }

  function finishProgress() {
    clearInterval(progressInterval);
    progressBar.style.width = '100%';
    setTimeout(resetProgress, 1200);
  }

  function resetProgress() {
    clearInterval(progressInterval);
    progressWrap.style.display = 'none';
    progressBar.style.width = '0%';
    progressVal = 0;
  }

  // ── UI HELPERS ────────────────────────────────────
  function setLoading(on) {
    convertBtn.disabled = on;
    btnLabel.style.display = on ? 'none' : 'flex';
    btnSpinner.style.display = on ? 'flex' : 'none';
  }

  function showError(msg) {
    errorMsg.textContent = msg;
    alertError.style.display = 'flex';
    alertSuccess.style.display = 'none';
  }

  function hideAlerts() {
    alertError.style.display = 'none';
    alertSuccess.style.display = 'none';
  }
})();
