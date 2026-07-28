/**
 * D-Office PCVT - Main Application Logic
 * ========================================
 * Quản lý văn bản đến - Chỉ đạo - Chuyển chỉ đạo
 * Dữ liệu lưu trên Google Sheets qua Apps Script API
 */

// ============================================================
// CONFIGURATION
// ============================================================

// ⚠️ REPLACE THIS with your deployed Google Apps Script URL
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwiyeYaCCX9yA8VgjXWzxBDdExIqsBMPZOoni-s2Zd5ujjyQmVaqIsJRDTZogX_zNE/exec';

// Path to documents.json (relative to app folder)
const DOCS_JSON_PATH = '../documents.json';

// Path to extracted files
const EXTRACTED_BASE = '../SOURCE D-O/extracted/';

// ============================================================
// ORGANIZATIONAL STRUCTURE DATA
// ============================================================

const ORG_STRUCTURE = {
  phongBanDonVi: [
    'Văn phòng',
    'Phòng Tổ chức & Nhân sự',
    'Phòng Kế hoạch & Vật Tư',
    'Phòng Quản lý đầu tư',
    'Phòng Kỹ thuật & An Toàn',
    'Phòng Kinh doanh',
    'Phòng Tài chính Kế toán',
    'Đội Vận hành lưới điện',
    'Đội Quản lý lưới điện',
    'Đội Dịch vụ khách hàng',
    'Đội Quản lý thu ghi',
    'Đội Quản lý hệ thống đo đếm',
    'Điện lực Đặc khu Côn Đảo',
  ],
  dangUyDoanThe: [
    'BCH Đảng ủy',
    'Nhóm Đảng vụ',
    'Chi bộ 1',
    'Chi bộ 2',
    'Chi bộ 3',
    'Chi bộ 4',
    'Chi bộ 5',
    'Chi bộ 6',
    'Công đoàn',
    'Đoàn thanh niên',
  ],
  toCongTac: [
    'Tổ Thẩm định',
    'Tổ Chuyên gia',
  ],
  banGiamDocCCD: [
    'GĐ - Nguyễn Ngọc Tuyền',
    'PGĐ KT - Trần Thanh Hải',
    'PGĐ KD - Đặng Quang Trung',
    'PGĐ ĐTXD - Lại Xuân Phương',
  ],
};

// ============================================================
// APPLICATION STATE
// ============================================================

let allDocuments = [];
let filteredDocuments = [];
let selectedDoc = null;
let chiDaoData = {};     // vanBanId -> array of chi dao
let chuyenChiDaoData = {}; // vanBanId -> array
let driveFileMapping = null; // folderName/fileName -> Google Drive preview URL

// ============================================================
// INITIALIZATION
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
  initUI();
  loadDocuments();
  loadDriveMapping();
  populateDynamicForms();
});

function initUI() {
  // Search
  document.getElementById('searchInput').addEventListener('input', (e) => {
    filterDocuments(e.target.value.trim());
  });

  // Tabs
  document.getElementById('tabChuTri').addEventListener('click', () => switchTab('chutri'));
  document.getElementById('tabPhoiHop').addEventListener('click', () => switchTab('phoihop'));

  // Action buttons
  document.getElementById('btnChiDao').addEventListener('click', () => openChiDaoModal());
  document.getElementById('btnChuyenChiDao').addEventListener('click', () => openChuyenChiDaoModal());
  document.getElementById('btnThongTin').addEventListener('click', () => {
    openChiDaoModal();
    switchModalTab('thongtin');
  });

  // Modal close buttons
  document.getElementById('closeChiDao').addEventListener('click', closeChiDaoModal);
  document.getElementById('closeChuyenChiDao').addEventListener('click', closeChuyenChiDaoModal);

  // Modal tab switching
  document.getElementById('tabChiDaoForm').addEventListener('click', () => switchModalTab('chidao'));
  document.getElementById('tabThongTinVB').addEventListener('click', () => switchModalTab('thongtin'));

  // Save buttons
  document.getElementById('btnSaveChiDao').addEventListener('click', saveChiDao);
  document.getElementById('btnSaveChuyenChiDao').addEventListener('click', saveChuyenChiDao);
  document.getElementById('btnCCDBack').addEventListener('click', closeChuyenChiDaoModal);

  // Close modal on overlay click
  document.getElementById('modalChiDao').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeChiDaoModal();
  });
  document.getElementById('modalChuyenChiDao').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeChuyenChiDaoModal();
  });

  // Keyboard shortcuts
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeChiDaoModal();
      closeChuyenChiDaoModal();
    }
  });

  // Preview panel toggle (collapse/expand file preview in chi dao modal)
  document.getElementById('previewToggleBtn').addEventListener('click', () => {
    const panel = document.getElementById('modalFilePreview');
    const icon = document.querySelector('#previewToggleBtn .material-icons-outlined');
    panel.classList.toggle('collapsed');
    if (panel.classList.contains('collapsed')) {
      icon.textContent = 'chevron_right';
    } else {
      icon.textContent = 'chevron_left';
    }
  });

  // PB group toggle
  document.querySelectorAll('.pb-group-header[data-toggle]').forEach(header => {
    header.addEventListener('click', (e) => {
      // Don't toggle if clicking on radio/checkbox
      if (e.target.tagName === 'INPUT') return;
      const body = header.nextElementSibling;
      if (body) {
        body.classList.toggle('collapsed');
        header.classList.toggle('collapsed');
      }
    });
  });

  // FAB
  document.getElementById('fabBtn').addEventListener('click', () => {
    if (selectedDoc) openChiDaoModal();
  });

  // Set default date (1 month from now)
  const defaultDate = new Date();
  defaultDate.setMonth(defaultDate.getMonth() + 1);
  document.getElementById('cdHanGiaiQuyet').value = defaultDate.toISOString().split('T')[0];
}

// ============================================================
// DATA LOADING
// ============================================================

async function loadDocuments() {
  const loading = document.getElementById('docListLoading');
  loading.style.display = 'flex';

  try {
    // Try Google Sheets first, fallback to local JSON
    if (GAS_URL) {
      try {
        const resp = await fetch(`${GAS_URL}?action=getDocuments`);
        const result = await resp.json();
        if (result.success && result.data.length > 0) {
          allDocuments = result.data;
          // Also try to load chi dao data
          loadChiDaoFromSheets();
        } else {
          await loadFromJSON();
        }
      } catch (e) {
        console.warn('GAS fetch failed, falling back to local JSON:', e);
        await loadFromJSON();
      }
    } else {
      await loadFromJSON();
    }

    filteredDocuments = [...allDocuments];
    renderDocumentList();
    updateCounts();
  } catch (e) {
    console.error('Failed to load documents:', e);
    loading.innerHTML = '<p style="color:var(--danger)">Lỗi tải dữ liệu</p>';
  }

  loading.style.display = 'none';
}

async function loadFromJSON() {
  const resp = await fetch(DOCS_JSON_PATH);
  allDocuments = await resp.json();
}

async function loadChiDaoFromSheets() {
  if (!GAS_URL) return;
  try {
    const [cdResp, ccdResp] = await Promise.all([
      fetch(`${GAS_URL}?action=getAllChiDao`),
      fetch(`${GAS_URL}?action=getAllChuyenChiDao`),
    ]);
    const cdResult = await cdResp.json();
    const ccdResult = await ccdResp.json();
    
    if (cdResult.success) {
      cdResult.data.forEach(cd => {
        if (!chiDaoData[cd.vanBanId]) chiDaoData[cd.vanBanId] = [];
        chiDaoData[cd.vanBanId].push(cd);
      });
    }
    if (ccdResult.success) {
      ccdResult.data.forEach(ccd => {
        if (!chuyenChiDaoData[ccd.vanBanId]) chuyenChiDaoData[ccd.vanBanId] = [];
        chuyenChiDaoData[ccd.vanBanId].push(ccd);
      });
    }
  } catch (e) {
    console.warn('Failed to load chi dao data:', e);
  }
}

// ============================================================
// GOOGLE DRIVE FILE MAPPING
// ============================================================

async function loadDriveMapping() {
  if (!GAS_URL) return;
  
  try {
    const resp = await fetch(`${GAS_URL}?action=getDriveMapping`);
    const result = await resp.json();
    if (result.success && result.data) {
      driveFileMapping = result.data;
      console.log(`Drive mapping loaded: ${Object.keys(driveFileMapping).length} files`);
    }
  } catch (e) {
    console.warn('Failed to load drive mapping:', e);
  }
}

/**
 * Get the URL to view a PDF file.
 * Uses Google Drive preview URL if mapping exists, otherwise falls back to local path.
 */
function getPdfViewUrl(folderName, fileName) {
  // Try Google Drive mapping first
  if (driveFileMapping) {
    const key = folderName + '/' + fileName;
    if (driveFileMapping[key]) {
      return driveFileMapping[key];
    }
  }
  // Fallback to local path
  return `${EXTRACTED_BASE}${folderName}/${encodeURIComponent(fileName)}#toolbar=0`;
}
// ============================================================
// DOCUMENT LIST RENDERING
// ============================================================

function renderDocumentList() {
  const container = document.getElementById('docList');
  const loading = document.getElementById('docListLoading');
  
  // Clear existing cards (not loading)
  container.querySelectorAll('.doc-card').forEach(c => c.remove());

  if (filteredDocuments.length === 0) {
    const empty = document.createElement('div');
    empty.className = 'doc-list-loading';
    empty.innerHTML = '<p>Không tìm thấy văn bản</p>';
    container.appendChild(empty);
    return;
  }

  filteredDocuments.forEach((doc, idx) => {
    const card = createDocCard(doc, idx);
    container.appendChild(card);
  });
}

function createDocCard(doc, index) {
  const card = document.createElement('div');
  card.className = 'doc-card';
  card.dataset.id = doc.id;

  // Flag icon based on độ khẩn
  let flagHtml = '';
  let khanHtml = '';
  if (doc.doKhan === 'Khẩn') {
    flagHtml = '<span class="material-icons-outlined doc-flag-icon khan">flag</span>';
    khanHtml = '<span class="doc-card-khan khan">Khẩn</span>';
  } else if (doc.doKhan === 'Thượng khẩn') {
    flagHtml = '<span class="material-icons-outlined doc-flag-icon thuong-khan">flag</span>';
    khanHtml = '<span class="doc-card-khan thuong-khan">Thượng khẩn</span>';
  } else if (doc.doKhan === 'Hỏa tốc') {
    flagHtml = '<span class="material-icons-outlined doc-flag-icon hoa-toc">priority_high</span>';
    khanHtml = '<span class="doc-card-khan hoa-toc">Hỏa tốc</span>';
  }

  // Status indicator
  const hasChiDao = chiDaoData[doc.id] && chiDaoData[doc.id].length > 0;
  const statusDot = hasChiDao 
    ? '<span class="status-dot done" title="Đã chỉ đạo"></span>' 
    : '<span class="status-dot pending" title="Chờ xử lý"></span>';

  card.innerHTML = `
    <div class="doc-card-header">
      <div class="doc-card-flag">
        ${flagHtml}
        <span class="doc-card-so">${escapeHtml(doc.soVanBan)}</span>
        ${khanHtml}
      </div>
      <span class="doc-card-date">${formatDate(doc.ngayVanBan)}</span>
    </div>
    <div class="doc-card-org">
      <span>${escapeHtml(doc.coQuanBanHanh || '—')}</span>
    </div>
    <div class="doc-card-summary">${escapeHtml(doc.trichYeu || 'Không có trích yếu')}</div>
    <div class="doc-card-status">${statusDot}</div>
  `;

  card.addEventListener('click', () => selectDocument(doc));

  return card;
}

function selectDocument(doc) {
  selectedDoc = doc;

  // Highlight active card
  document.querySelectorAll('.doc-card').forEach(c => c.classList.remove('active'));
  const activeCard = document.querySelector(`.doc-card[data-id="${doc.id}"]`);
  if (activeCard) activeCard.classList.add('active');

  // Show toolbar actions
  document.getElementById('toolbarActions').style.display = 'flex';

  // Update preview
  renderPreview(doc);
}

function renderPreview(doc) {
  document.getElementById('previewPlaceholder').style.display = 'none';
  const content = document.getElementById('previewContent');
  content.style.display = 'flex';

  // File links
  const linksContainer = document.getElementById('previewFileLinks');
  linksContainer.innerHTML = '';
  if (doc.files && doc.files.length > 0) {
    doc.files.forEach(file => {
      const ext = file.split('.').pop().toLowerCase();
      const icon = ext === 'pdf' ? 'picture_as_pdf' 
        : ext === 'docx' || ext === 'doc' ? 'description' 
        : ext === 'xlsx' ? 'table_chart' 
        : 'insert_drive_file';
      const link = document.createElement('a');
      link.className = 'file-link';
      link.href = `${EXTRACTED_BASE}${doc.folderName}/${file}`;
      link.target = '_blank';
      link.innerHTML = `<span class="material-icons-outlined">${icon}</span>${truncateFilename(file, 30)}`;
      linksContainer.appendChild(link);
    });
  }

  // Show first PDF in preview area
  const previewArea = document.getElementById('previewFileArea');
  const pdfFile = doc.files ? doc.files.find(f => f.toLowerCase().endsWith('.pdf')) : null;
  if (pdfFile) {
    const pdfUrl = getPdfViewUrl(doc.folderName, pdfFile);
    previewArea.innerHTML = `<iframe src="${pdfUrl}" title="PDF Preview"></iframe>`;
  } else {
    previewArea.innerHTML = '<p class="preview-msg">Không có file PDF để xem trước</p>';
  }
}

// ============================================================
// FILTER & SEARCH
// ============================================================

function filterDocuments(query) {
  if (!query) {
    filteredDocuments = [...allDocuments];
  } else {
    const q = query.toLowerCase();
    filteredDocuments = allDocuments.filter(doc => {
      return (doc.soVanBan && doc.soVanBan.toLowerCase().includes(q))
        || (doc.trichYeu && doc.trichYeu.toLowerCase().includes(q))
        || (doc.coQuanBanHanh && doc.coQuanBanHanh.toLowerCase().includes(q));
    });
  }
  renderDocumentList();
  updateCounts();
}

function switchTab(tab) {
  document.getElementById('tabChuTri').classList.toggle('active', tab === 'chutri');
  document.getElementById('tabPhoiHop').classList.toggle('active', tab === 'phoihop');
}

function updateCounts() {
  document.getElementById('countChuTri').textContent = filteredDocuments.length;
  document.getElementById('countPhoiHop').textContent = '0';
}

// ============================================================
// CHỈ ĐẠO MODAL
// ============================================================

function openChiDaoModal() {
  if (!selectedDoc) {
    showToast('Vui lòng chọn một văn bản', 'error');
    return;
  }

  const modal = document.getElementById('modalChiDao');
  modal.classList.add('show');
  document.body.style.overflow = 'hidden';

  // Ensure preview panel is expanded
  const previewPanel = document.getElementById('modalFilePreview');
  previewPanel.classList.remove('collapsed');
  const toggleIcon = document.querySelector('#previewToggleBtn .material-icons-outlined');
  if (toggleIcon) toggleIcon.textContent = 'chevron_left';

  // Set document title
  document.getElementById('cdDocTitle').textContent = selectedDoc.soVanBan;

  // File links in modal
  const linksContainer = document.getElementById('modalFileLinks');
  linksContainer.innerHTML = '';
  if (selectedDoc.files && selectedDoc.files.length > 0) {
    selectedDoc.files.forEach(file => {
      const ext = file.split('.').pop().toLowerCase();
      const icon = ext === 'pdf' ? 'picture_as_pdf' : 'description';
      const link = document.createElement('a');
      link.className = 'file-link';
      link.href = `${EXTRACTED_BASE}${selectedDoc.folderName}/${file}`;
      link.target = '_blank';
      link.innerHTML = `<span class="material-icons-outlined">${icon}</span>${truncateFilename(file, 25)}`;
      link.addEventListener('click', (e) => {
        e.preventDefault();
        // Load in modal viewer
        if (file.toLowerCase().endsWith('.pdf')) {
          const viewer = document.getElementById('modalFileViewer');
          viewer.innerHTML = `<iframe src="${getPdfViewUrl(selectedDoc.folderName, file)}" title="PDF"></iframe>`;
        } else {
          window.open(link.href, '_blank');
        }
      });
      linksContainer.appendChild(link);
    });

    // Auto-load first PDF
    const pdfFile = selectedDoc.files.find(f => f.toLowerCase().endsWith('.pdf'));
    if (pdfFile) {
      document.getElementById('modalFileViewer').innerHTML = 
        `<iframe src="${getPdfViewUrl(selectedDoc.folderName, pdfFile)}" title="PDF"></iframe>`;
    }
  }

  // Populate thông tin văn bản tab
  document.getElementById('infoTrichYeu').textContent = selectedDoc.trichYeu || '—';
  document.getElementById('infoSoVB').textContent = selectedDoc.soVanBan;
  document.getElementById('infoNgayVB').textContent = formatDate(selectedDoc.ngayVanBan);
  document.getElementById('infoHanGQ').textContent = '—';
  document.getElementById('infoCQBH').textContent = selectedDoc.coQuanBanHanh || '—';
  document.getElementById('infoDoKhan').textContent = selectedDoc.doKhan || 'Bình thường';
  document.getElementById('infoDoMat').textContent = selectedDoc.doMat || 'Bình thường';

  // File info
  const infoFiles = document.getElementById('infoFiles');
  infoFiles.innerHTML = '';
  if (selectedDoc.files && selectedDoc.files.length > 0) {
    selectedDoc.files.forEach(file => {
      const ext = file.split('.').pop().toLowerCase();
      const icon = ext === 'pdf' ? 'picture_as_pdf' : 'description';
      const div = document.createElement('div');
      div.className = 'file-download';
      div.innerHTML = `<span class="material-icons-outlined">${icon}</span>${file}<br><a href="${EXTRACTED_BASE}${selectedDoc.folderName}/${file}" target="_blank" style="color:var(--primary);font-size:0.78rem">Tải xuống</a>`;
      infoFiles.appendChild(div);
    });
  } else {
    infoFiles.textContent = '—';
  }

  // Load existing chi dao if any
  loadExistingChiDao();

  // Default to chi dao tab
  switchModalTab('chidao');
}

function closeChiDaoModal() {
  document.getElementById('modalChiDao').classList.remove('show');
  document.body.style.overflow = '';
}

function switchModalTab(tab) {
  document.getElementById('tabChiDaoForm').classList.toggle('active', tab === 'chidao');
  document.getElementById('tabThongTinVB').classList.toggle('active', tab === 'thongtin');
  document.getElementById('contentChiDao').classList.toggle('active', tab === 'chidao');
  document.getElementById('contentThongTin').classList.toggle('active', tab === 'thongtin');
}

function loadExistingChiDao() {
  if (!selectedDoc) return;
  const existing = chiDaoData[selectedDoc.id];
  if (existing && existing.length > 0) {
    const latest = existing[existing.length - 1];
    document.getElementById('cdNoiDung').value = latest.noiDung || '';
    
    // Restore radio selection for chủ trì
    if (latest.chuTri) {
      const radios = document.querySelectorAll('input[name="cd_chutri"]');
      radios.forEach(r => {
        r.checked = r.value === latest.chuTri;
      });
    }

    // Restore checkboxes for phối hợp
    if (latest.phoiHop) {
      const phList = latest.phoiHop.split(',').map(s => s.trim());
      document.querySelectorAll('input[data-group="phoihop"]').forEach(cb => {
        cb.checked = phList.includes(cb.value);
      });
    }

    // Restore checkboxes for xem để biết
    if (latest.xemDeBiet) {
      const xdbList = latest.xemDeBiet.split(',').map(s => s.trim());
      document.querySelectorAll('input[data-group="xemdebiet"]').forEach(cb => {
        cb.checked = xdbList.includes(cb.value);
      });
    }

    if (latest.hanGiaiQuyet) {
      document.getElementById('cdHanGiaiQuyet').value = latest.hanGiaiQuyet;
    }
    if (latest.doKhan) {
      document.getElementById('cdDoKhan').value = latest.doKhan;
    }
    document.getElementById('cdYeuCauTraLoi').checked = latest.yeuCauTraLoi === 'Có';
    document.getElementById('cdChuyenThuKy').checked = latest.chuyenThuKy === 'Có';
    document.getElementById('cdTheoDoiVB').checked = latest.theoDoiVanBan === 'Có';
  } else {
    // Reset form
    document.getElementById('cdNoiDung').value = '';
    document.querySelectorAll('input[name="cd_chutri"]').forEach(r => r.checked = false);
    document.querySelectorAll('input[data-group="phoihop"]').forEach(cb => cb.checked = false);
    document.querySelectorAll('input[data-group="xemdebiet"]').forEach(cb => cb.checked = false);
    document.getElementById('cdYeuCauTraLoi').checked = false;
    document.getElementById('cdChuyenThuKy').checked = false;
    document.getElementById('cdTheoDoiVB').checked = false;
  }
}

async function saveChiDao() {
  if (!selectedDoc) return;

  // Gather form data
  const noiDung = document.getElementById('cdNoiDung').value.trim();
  const hanGiaiQuyet = document.getElementById('cdHanGiaiQuyet').value;
  const doKhan = document.getElementById('cdDoKhan').value;
  const yeuCauTraLoi = document.getElementById('cdYeuCauTraLoi').checked;
  const chuyenThuKy = document.getElementById('cdChuyenThuKy').checked;
  const theoDoiVanBan = document.getElementById('cdTheoDoiVB').checked;

  // Get chủ trì (radio)
  const chuTriRadio = document.querySelector('input[name="cd_chutri"]:checked');
  const chuTri = chuTriRadio ? chuTriRadio.value : '';

  // Get phối hợp (checkboxes)
  const phoiHop = [];
  document.querySelectorAll('input[data-group="phoihop"]:checked').forEach(cb => {
    phoiHop.push(cb.value);
  });

  // Get xem để biết (checkboxes)
  const xemDeBiet = [];
  document.querySelectorAll('input[data-group="xemdebiet"]:checked').forEach(cb => {
    xemDeBiet.push(cb.value);
  });

  const data = {
    action: 'saveChiDao',
    vanBanId: selectedDoc.id,
    noiDung,
    chuTri,
    phoiHop: phoiHop.join(', '),
    xemDeBiet: xemDeBiet.join(', '),
    hanGiaiQuyet,
    doKhan,
    yeuCauTraLoi,
    chuyenThuKy,
    theoDoiVanBan,
    nguoiChiDao: 'Nguyễn Ngọc Tuyền',
  };

  // Save to Google Sheets
  if (GAS_URL) {
    try {
      const resp = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      const result = await resp.json();
      if (result.success) {
        showToast('Đã lưu chỉ đạo thành công!', 'success');
        // Update local cache
        if (!chiDaoData[selectedDoc.id]) chiDaoData[selectedDoc.id] = [];
        chiDaoData[selectedDoc.id].push(data);
        closeChiDaoModal();
        renderDocumentList(); // Update status dots
        return;
      } else {
        showToast('Lỗi: ' + (result.error || 'Không xác định'), 'error');
      }
    } catch (e) {
      console.error('Save chi dao error:', e);
      showToast('Lỗi kết nối server', 'error');
    }
  } else {
    // Save locally (demo mode)
    if (!chiDaoData[selectedDoc.id]) chiDaoData[selectedDoc.id] = [];
    chiDaoData[selectedDoc.id].push(data);
    showToast('Đã lưu chỉ đạo (chế độ offline)', 'success');
    closeChiDaoModal();
    renderDocumentList();
  }
}

// ============================================================
// CHUYỂN CHỈ ĐẠO MODAL
// ============================================================

function openChuyenChiDaoModal() {
  if (!selectedDoc) {
    showToast('Vui lòng chọn một văn bản', 'error');
    return;
  }

  document.getElementById('modalChuyenChiDao').classList.add('show');
  document.body.style.overflow = 'hidden';

  // Reset form
  document.getElementById('ccdNoiDung').value = '';
  document.querySelectorAll('#modalChuyenChiDao input[type="radio"]').forEach(r => r.checked = false);
  document.querySelectorAll('#modalChuyenChiDao input[type="checkbox"]').forEach(cb => cb.checked = false);
}

function closeChuyenChiDaoModal() {
  document.getElementById('modalChuyenChiDao').classList.remove('show');
  document.body.style.overflow = '';
}

async function saveChuyenChiDao() {
  if (!selectedDoc) return;

  const noiDungChuyen = document.getElementById('ccdNoiDung').value.trim();
  
  const chuTriRadio = document.querySelector('#modalChuyenChiDao input[name="ccd_chutri"]:checked');
  const chuTri = chuTriRadio ? chuTriRadio.value : '';

  const xemDeBiet = [];
  document.querySelectorAll('#modalChuyenChiDao input[data-group="ccd_xemdebiet"]:checked').forEach(cb => {
    xemDeBiet.push(cb.value);
  });

  const data = {
    action: 'saveChuyenChiDao',
    vanBanId: selectedDoc.id,
    noiDungChuyen,
    chuTri,
    xemDeBiet: xemDeBiet.join(', '),
    nguoiChuyen: 'Nguyễn Ngọc Tuyền',
  };

  if (GAS_URL) {
    try {
      const resp = await fetch(GAS_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      });
      const result = await resp.json();
      if (result.success) {
        showToast('Đã lưu chuyển chỉ đạo!', 'success');
        if (!chuyenChiDaoData[selectedDoc.id]) chuyenChiDaoData[selectedDoc.id] = [];
        chuyenChiDaoData[selectedDoc.id].push(data);
        closeChuyenChiDaoModal();
        return;
      } else {
        showToast('Lỗi: ' + (result.error || 'Không xác định'), 'error');
      }
    } catch (e) {
      showToast('Lỗi kết nối server', 'error');
    }
  } else {
    if (!chuyenChiDaoData[selectedDoc.id]) chuyenChiDaoData[selectedDoc.id] = [];
    chuyenChiDaoData[selectedDoc.id].push(data);
    showToast('Đã lưu chuyển chỉ đạo (offline)', 'success');
    closeChuyenChiDaoModal();
  }
}

// ============================================================
// POPULATE DYNAMIC FORM SECTIONS
// ============================================================

function populateDynamicForms() {
  // Phòng ban đơn vị
  const pbdvBody = document.getElementById('pb-pbdv');
  ORG_STRUCTURE.phongBanDonVi.forEach((name, i) => {
    pbdvBody.appendChild(createPBRow(name, i % 2 === 1));
  });

  // Đảng ủy & Đoàn thể
  const dudtBody = document.getElementById('pb-dudt');
  ORG_STRUCTURE.dangUyDoanThe.forEach((name, i) => {
    dudtBody.appendChild(createPBRow(name, i % 2 === 1));
  });

  // Tổ công tác
  const tctBody = document.getElementById('pb-tct');
  ORG_STRUCTURE.toCongTac.forEach((name, i) => {
    tctBody.appendChild(createPBRow(name, i % 2 === 1));
  });

  // Chuyển chỉ đạo - Ban giám đốc
  const ccdBGD = document.getElementById('ccd-bgd-body');
  ORG_STRUCTURE.banGiamDocCCD.forEach((name, i) => {
    ccdBGD.appendChild(createCCDRow(name, i % 2 === 1));
  });

  // Chuyển chỉ đạo - Đảng ủy & Đoàn thể
  const ccdDUDT = document.getElementById('ccd-dudt-body');
  ORG_STRUCTURE.dangUyDoanThe.forEach((name, i) => {
    ccdDUDT.appendChild(createCCDRow(name, i % 2 === 1));
  });
}

function createPBRow(name, isAlt) {
  const row = document.createElement('div');
  row.className = 'pb-row' + (isAlt ? ' alt' : '');
  row.innerHTML = `
    <span class="pb-name">${escapeHtml(name)}</span>
    <div class="pb-cols">
      <span class="pb-radio-cell">
        <input type="radio" name="cd_chutri" value="${escapeHtml(name)}">
      </span>
      <span class="pb-check-cell">
        <input type="checkbox" data-group="phoihop" value="${escapeHtml(name)}">
      </span>
      <span class="pb-check-cell">
        <input type="checkbox" data-group="xemdebiet" value="${escapeHtml(name)}">
      </span>
    </div>
  `;
  return row;
}

function createCCDRow(name, isAlt) {
  const row = document.createElement('div');
  row.className = 'ccd-row' + (isAlt ? ' alt' : '');
  row.innerHTML = `
    <span class="ccd-name">${escapeHtml(name)}</span>
    <span class="ccd-radio-cell">
      <input type="radio" name="ccd_chutri" value="${escapeHtml(name)}">
    </span>
    <span class="ccd-check-cell">
      <input type="checkbox" data-group="ccd_xemdebiet" value="${escapeHtml(name)}">
    </span>
  `;
  return row;
}

// ============================================================
// TOAST NOTIFICATIONS
// ============================================================

function showToast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  
  const icon = type === 'success' ? 'check_circle' 
    : type === 'error' ? 'error' 
    : 'info';
  
  toast.innerHTML = `<span class="material-icons-outlined">${icon}</span>${escapeHtml(message)}`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.remove();
  }, 3000);
}

// ============================================================
// UTILITIES
// ============================================================

function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

function truncateFilename(name, maxLen) {
  if (name.length <= maxLen) return name;
  const ext = name.split('.').pop();
  const base = name.slice(0, maxLen - ext.length - 4);
  return base + '....' + ext;
}

/**
 * Format date from various formats (DD/MM/YYYY, ISO string, etc.)
 * Always returns DD/MM/YYYY format.
 */
function formatDate(dateStr) {
  if (!dateStr) return '—';
  
  // Already DD/MM/YYYY format
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
    return dateStr;
  }
  
  // ISO format (2026-07-26T17:00:00.000Z) or similar
  try {
    const d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      const day = String(d.getDate()).padStart(2, '0');
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const year = d.getFullYear();
      return `${day}/${month}/${year}`;
    }
  } catch (e) {}
  
  return String(dateStr);
}
