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
    'GĐ - Nguyễn Ngọc Tuyến',
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
let suggestedDocs = {}; // vanBanId -> { pgd, chuTri, phoiHop, xemDeBiet } - AI suggestions

// ============================================================
// AI AUTO-SUGGEST RULES
// ============================================================

const SUGGEST_RULES = [
  { keywords: ['thi công', 'hotline', 'liveline', 'phương án thi công', 'sự cố', 'đường dây', 'mang điện'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Phòng Kỹ thuật & An Toàn'], phoiHop: ['Đội Vận hành lưới điện'] },
  { keywords: ['vận hành', 'điều độ', 'phương thức vận hành', 'hệ thống điện'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Phòng Kỹ thuật & An Toàn'], phoiHop: ['Đội Vận hành lưới điện'] },
  { keywords: ['điện kế', 'đo đếm', 'công tơ', 'cấp điện kế'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Đội Quản lý hệ thống đo đếm'], phoiHop: ['Phòng Kỹ thuật & An Toàn'] },
  { keywords: ['năng lượng mặt trời', 'mặt trời mái nhà', 'NLMT', 'điện mặt trời'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Phòng Kinh doanh'], phoiHop: ['Đội Dịch vụ khách hàng'] },
  { keywords: ['tuyển dụng', 'nhân sự', 'vị trí chức danh', 'khung năng lực', 'cơ cấu tổ chức'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Phòng Tổ chức & Nhân sự'], phoiHop: ['Văn phòng'] },
  { keywords: ['đào tạo', 'thạc sĩ', 'tuyển sinh', 'sát hạch', 'nghiệp vụ'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Phòng Tổ chức & Nhân sự'], phoiHop: ['Văn phòng'] },
  { keywords: ['phúc lợi', 'kỷ niệm', 'lễ', '50 năm', 'sự kiện'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Phòng Tổ chức & Nhân sự'], phoiHop: ['Văn phòng'] },
  { keywords: ['đầu tư', 'xây dựng', 'tái định cư', 'trạm biến áp', 'hạ tầng kỹ thuật'],
    pgd: 'PGĐ ĐTXD - Lại Xuân Phương', chuTri: ['Phòng Quản lý đầu tư'], phoiHop: ['Phòng Kỹ thuật & An Toàn'] },
  { keywords: ['pháp lý', 'trụ sở', 'kho bãi', 'phòng biến điện'],
    pgd: 'PGĐ ĐTXD - Lại Xuân Phương', chuTri: ['Phòng Quản lý đầu tư'], phoiHop: ['Văn phòng'] },
  { keywords: ['cưỡng chế', 'thu hồi đất', 'giải phóng mặt bằng'],
    pgd: 'PGĐ ĐTXD - Lại Xuân Phương', chuTri: ['Phòng Quản lý đầu tư'], phoiHop: ['Phòng Kỹ thuật & An Toàn', 'Văn phòng'] },
  { keywords: ['vật tư', 'vận chuyển', 'phụ tùng', 'thiết bị'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Phòng Kế hoạch & Vật Tư'], phoiHop: ['Phòng Kỹ thuật & An Toàn'] },
  { keywords: ['kinh doanh', 'khách hàng', 'giá điện', 'tiền điện'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Phòng Kinh doanh'], phoiHop: ['Đội Dịch vụ khách hàng'] },
  { keywords: ['tài chính', 'kế toán', 'chi phí', 'ngân sách'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Phòng Tài chính Kế toán'], phoiHop: ['Văn phòng'] },
  { keywords: ['đảng', 'nghị quyết', 'BCH', 'trung ương', 'đảng ủy', 'chi bộ'],
    pgd: '', chuTri: ['BCH Đảng ủy'], phoiHop: ['Văn phòng'] },
  { keywords: ['công nghệ thông tin', 'dữ liệu', 'AI', 'khoa học dữ liệu'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Văn phòng'], phoiHop: ['Phòng Kỹ thuật & An Toàn'] },
  { keywords: ['luật', 'pháp luật', 'an ninh dữ liệu', 'dự án luật', 'góp ý'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Văn phòng'], phoiHop: ['Phòng Kỹ thuật & An Toàn'] },
  { keywords: ['giấy mời', 'hội nghị', 'họp'],
    pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Văn phòng'], phoiHop: [] },
  { keywords: ['công suất phản kháng', 'ranh giới', 'đo ranh'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Phòng Kỹ thuật & An Toàn'], phoiHop: ['Đội Vận hành lưới điện'] },
  { keywords: ['lưới điện', 'quản lý lưới'],
    pgd: 'PGĐ KT - Trần Thanh Hải', chuTri: ['Đội Quản lý lưới điện'], phoiHop: ['Phòng Kỹ thuật & An Toàn'] },
];

function suggestAssignment(doc) {
  const text = ((doc.trichYeu || '') + ' ' + (doc.coQuanBanHanh || '')).toLowerCase();
  
  for (const rule of SUGGEST_RULES) {
    for (const kw of rule.keywords) {
      if (text.includes(kw.toLowerCase())) {
        return {
          pgd: rule.pgd,
          chuTri: rule.chuTri || [],
          phoiHop: rule.phoiHop || [],
          xemDeBiet: [],
        };
      }
    }
  }
  // Default
  return { pgd: 'PGĐ KD - Đặng Quang Trung', chuTri: ['Văn phòng'], phoiHop: [], xemDeBiet: [] };
}

function applySuggestion() {
  if (allDocuments.length === 0) {
    showToast('Chưa có văn bản nào để đề xuất', 'error');
    return;
  }
  
  let count = 0;
  allDocuments.forEach(doc => {
    // Skip docs that already have chi dao saved
    if (chiDaoData[doc.id] && chiDaoData[doc.id].length > 0) return;
    
    const s = suggestAssignment(doc);
    suggestedDocs[doc.id] = s;
    count++;
  });
  
  showToast(`🤖 Đã đề xuất phân công cho ${count} văn bản`, 'info');
  renderDocumentList(); // Refresh dots
  
  // If chi dao modal is open with a selected doc, apply to checkboxes too
  if (selectedDoc && suggestedDocs[selectedDoc.id]) {
    const s = suggestedDocs[selectedDoc.id];
    document.querySelectorAll('#modalChiDao input[data-group="chutri"]').forEach(cb => cb.checked = false);
    document.querySelectorAll('#modalChiDao input[data-group="phoihop"]').forEach(cb => cb.checked = false);
    document.querySelectorAll('#modalChiDao input[data-group="xemdebiet"]').forEach(cb => cb.checked = false);
    
    s.chuTri.forEach(name => {
      const cb = document.querySelector(`#modalChiDao input[data-group="chutri"][value="${name}"]`);
      if (cb) { cb.checked = true; cb.closest('.pb-row')?.classList.add('suggested'); }
    });
    s.phoiHop.forEach(name => {
      const cb = document.querySelector(`#modalChiDao input[data-group="phoihop"][value="${name}"]`);
      if (cb) { cb.checked = true; cb.closest('.pb-row')?.classList.add('suggested'); }
    });
    document.querySelectorAll('#modalChiDao .pb-row').forEach(row => {
      const anyChecked = row.querySelectorAll('input:checked').length > 0;
      row.classList.toggle('suggested', anyChecked);
    });
  }
}

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
  document.getElementById('tabFilter').addEventListener('click', (e) => { e.stopPropagation(); toggleFilterDropdown(); });

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

  // Suggest button
  const btnSuggest = document.getElementById('btnSuggestAI');
  if (btnSuggest) btnSuggest.addEventListener('click', applySuggestion);

  // Close modal on overlay click
  document.getElementById('modalChiDao').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeChiDaoModal();
  });
  document.getElementById('modalChuyenChiDao').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeChuyenChiDaoModal();
  });
  document.getElementById('modalOverview').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeOverviewModal();
  });

  // Overview button
  document.getElementById('btnOverview').addEventListener('click', openOverviewModal);
  document.getElementById('closeOverview').addEventListener('click', closeOverviewModal);

  // Keyboard shortcuts
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeChiDaoModal();
      closeChuyenChiDaoModal();
      closeOverviewModal();
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
        const resp = await fetch(`${GAS_URL}?action=getDocuments`, { redirect: 'follow' });
        const text = await resp.text();
        const result = JSON.parse(text);
        if (result.success && result.data && result.data.length > 0) {
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
      fetch(`${GAS_URL}?action=getAllChiDao`, { redirect: 'follow' }),
      fetch(`${GAS_URL}?action=getAllChuyenChiDao`, { redirect: 'follow' }),
    ]);
    const cdText = await cdResp.text();
    const ccdText = await ccdResp.text();
    let cdResult, ccdResult;
    try { cdResult = JSON.parse(cdText); } catch(e) { cdResult = { success: false }; }
    try { ccdResult = JSON.parse(ccdText); } catch(e) { ccdResult = { success: false }; }
    
    if (cdResult.success && cdResult.data) {
      chiDaoData = {};
      cdResult.data.forEach(cd => {
        if (!chiDaoData[cd.vanBanId]) chiDaoData[cd.vanBanId] = [];
        chiDaoData[cd.vanBanId].push(cd);
      });
    }
    if (ccdResult.success && ccdResult.data) {
      chuyenChiDaoData = {};
      ccdResult.data.forEach(ccd => {
        if (!chuyenChiDaoData[ccd.vanBanId]) chuyenChiDaoData[ccd.vanBanId] = [];
        chuyenChiDaoData[ccd.vanBanId].push(ccd);
      });
    }
    // Re-render to update status dots
    renderDocumentList();
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
    const resp = await fetch(`${GAS_URL}?action=getDriveMapping`, {
      redirect: 'follow',
    });
    const text = await resp.text();
    try {
      const result = JSON.parse(text);
      if (result.success && result.data) {
        driveFileMapping = result.data;
        console.log(`Drive mapping loaded: ${Object.keys(driveFileMapping).length} files`);
      }
    } catch (parseErr) {
      console.warn('Drive mapping response is not JSON:', text.substring(0, 200));
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
  if (driveFileMapping) {
    // Key 1: folderName/fileName (subfolder structure)
    const key1 = folderName + '/' + fileName;
    if (driveFileMapping[key1]) return driveFileMapping[key1];
    
    // Key 2: fileName directly
    if (driveFileMapping[fileName]) return driveFileMapping[fileName];
    
    // Key 3: baseName (without .pdf)
    const baseName = fileName.replace(/\.pdf$/i, '');
    const baseKey = baseName + '/' + fileName;
    if (driveFileMapping[baseKey]) return driveFileMapping[baseKey];
    
    // Key 4: Search all mapping keys for exact fileName match
    for (const [mapKey, url] of Object.entries(driveFileMapping)) {
      if (mapKey.endsWith('/' + fileName) || mapKey === fileName) {
        return url;
      }
    }
    
    // Key 5: Fuzzy match - strip .signed suffixes and compare core name
    const coreFileName = fileName.replace(/\.pdf$/i, '').replace(/\.signed/g, '').toLowerCase().trim();
    for (const [mapKey, url] of Object.entries(driveFileMapping)) {
      const mapFile = mapKey.split('/').pop() || mapKey;
      const coreMapFile = mapFile.replace(/\.pdf$/i, '').replace(/\.signed/g, '').toLowerCase().trim();
      if (coreFileName === coreMapFile) {
        return url;
      }
    }
    
    // Key 6: Contains match - if core part of filename appears in mapping key
    for (const [mapKey, url] of Object.entries(driveFileMapping)) {
      if (mapKey.toLowerCase().includes(coreFileName) || coreFileName.includes(mapKey.split('/').pop().replace(/\.pdf$/i, '').replace(/\.signed/g, '').toLowerCase().trim())) {
        return url;
      }
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
  const isSuggested = suggestedDocs[doc.id] && !hasChiDao;
  let statusDot;
  if (hasChiDao) {
    statusDot = '<span class="status-dot done" title="Đã chỉ đạo"></span>';
  } else if (isSuggested) {
    statusDot = '<span class="status-dot suggested" title="Đề xuất AI"></span>';
  } else {
    statusDot = '<span class="status-dot pending" title="Chờ xử lý"></span>';
  }

  card.innerHTML = `
    <div class="doc-card-header">
      <div class="doc-card-flag">
        ${flagHtml}
        <span class="doc-card-so">${escapeHtml(doc.soVanBan)}</span>
        ${khanHtml}
      </div>
      <div class="doc-card-date-status">
        <span class="doc-card-date">${formatDate(doc.ngayVanBan)}</span>
        ${statusDot}
      </div>
    </div>
    <div class="doc-card-org">
      <span>${escapeHtml(doc.coQuanBanHanh || '—')}</span>
    </div>
    <div class="doc-card-summary">${escapeHtml(doc.trichYeu || 'Không có trích yếu')}</div>
    ${doc.ngayTaiLen ? `<div class="doc-card-upload-time"><span class="material-icons-outlined" style="font-size:12px">cloud_upload</span> ${escapeHtml(doc.ngayTaiLen)}</div>` : ''}
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

  // Update assignment summary panel
  updateAssignmentSummary(doc);
}

function updateAssignmentSummary(doc) {
  const panel = document.getElementById('assignmentSummary');
  const body = document.getElementById('assignmentBody');
  
  const cd = chiDaoData[doc.id] && chiDaoData[doc.id].length > 0 ? chiDaoData[doc.id][0] : null;
  const suggested = suggestedDocs[doc.id] || null;
  const source = cd || suggested;
  
  if (!source) {
    panel.style.display = 'none';
    return;
  }
  
  panel.style.display = 'block';
  const isSuggested = !cd && suggested;
  
  let html = '';
  
  // Chủ trì
  const chuTri = source.chuTri ? (Array.isArray(source.chuTri) ? source.chuTri.join(', ') : source.chuTri) : '';
  if (chuTri) {
    html += `<div class="assign-item"><span class="assign-label ct">Chủ trì:</span> <span class="assign-value">${escapeHtml(chuTri)}</span></div>`;
  }
  
  // Phối hợp
  const phoiHop = source.phoiHop ? (Array.isArray(source.phoiHop) ? source.phoiHop.join(', ') : source.phoiHop) : '';
  if (phoiHop) {
    html += `<div class="assign-item"><span class="assign-label ph">Phối hợp:</span> <span class="assign-value">${escapeHtml(phoiHop)}</span></div>`;
  }
  
  // Xem để biết
  const xdb = source.xemDeBiet ? (Array.isArray(source.xemDeBiet) ? source.xemDeBiet.join(', ') : source.xemDeBiet) : '';
  if (xdb) {
    html += `<div class="assign-item"><span class="assign-label xdb">Xem để biết:</span> <span class="assign-value">${escapeHtml(xdb)}</span></div>`;
  }
  
  if (!html) {
    panel.style.display = 'none';
    return;
  }
  
  if (isSuggested) {
    html = '<div class="assign-badge suggested">🤖 Đề xuất AI</div>' + html;
  } else {
    html = '<div class="assign-badge saved">✅ Đã phân công</div>' + html;
  }
  
  body.innerHTML = html;
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

let currentFilterMode = 'all'; // 'all', 'assigned', 'unassigned'

function switchTab(tab) {
  document.getElementById('tabChuTri').classList.toggle('active', tab === 'chutri');
  document.getElementById('tabPhoiHop').classList.toggle('active', tab === 'phoihop');
}

function toggleFilterDropdown() {
  let dropdown = document.getElementById('filterDropdown');
  if (dropdown) {
    dropdown.remove();
    return;
  }
  const btn = document.getElementById('tabFilter');
  const rect = btn.getBoundingClientRect();
  dropdown = document.createElement('div');
  dropdown.id = 'filterDropdown';
  dropdown.className = 'filter-dropdown';
  dropdown.innerHTML = `
    <div class="filter-option ${currentFilterMode === 'all' ? 'active' : ''}" data-filter="all">
      <span class="material-icons-outlined" style="font-size:16px">list</span> Tất cả
    </div>
    <div class="filter-option ${currentFilterMode === 'assigned' ? 'active' : ''}" data-filter="assigned">
      <span class="status-dot done" style="display:inline-block;vertical-align:middle;margin-right:6px"></span> Đã giao việc
    </div>
    <div class="filter-option ${currentFilterMode === 'unassigned' ? 'active' : ''}" data-filter="unassigned">
      <span class="status-dot pending" style="display:inline-block;vertical-align:middle;margin-right:6px"></span> Chưa giao việc
    </div>
  `;
  dropdown.style.position = 'absolute';
  dropdown.style.top = (rect.bottom + 4) + 'px';
  dropdown.style.left = (rect.left) + 'px';
  document.body.appendChild(dropdown);

  dropdown.querySelectorAll('.filter-option').forEach(opt => {
    opt.addEventListener('click', () => {
      currentFilterMode = opt.dataset.filter;
      applyFilterMode();
      dropdown.remove();
    });
  });

  // Close on outside click
  setTimeout(() => {
    document.addEventListener('click', function closeDropdown(e) {
      if (!dropdown.contains(e.target) && e.target !== btn) {
        dropdown.remove();
        document.removeEventListener('click', closeDropdown);
      }
    });
  }, 10);
}

function applyFilterMode() {
  const query = document.getElementById('searchInput') ? document.getElementById('searchInput').value.trim().toLowerCase() : '';
  let docs = [...allDocuments];
  
  // Apply search filter
  if (query) {
    docs = docs.filter(doc => {
      return (doc.soVanBan && doc.soVanBan.toLowerCase().includes(query))
        || (doc.trichYeu && doc.trichYeu.toLowerCase().includes(query))
        || (doc.coQuanBanHanh && doc.coQuanBanHanh.toLowerCase().includes(query));
    });
  }
  
  // Apply assignment filter
  if (currentFilterMode === 'assigned') {
    docs = docs.filter(doc => chiDaoData[doc.id] && chiDaoData[doc.id].length > 0);
  } else if (currentFilterMode === 'unassigned') {
    docs = docs.filter(doc => !chiDaoData[doc.id] || chiDaoData[doc.id].length === 0);
  }
  
  filteredDocuments = docs;
  renderDocumentList();
  updateCounts();
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
    
    // Restore checkboxes for chủ trì
    if (latest.chuTri) {
      const ctList = latest.chuTri.split(',').map(s => s.trim());
      document.querySelectorAll('input[data-group="chutri"]').forEach(cb => {
        cb.checked = ctList.includes(cb.value);
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
    document.querySelectorAll('input[data-group="chutri"]').forEach(cb => cb.checked = false);
    document.querySelectorAll('input[data-group="phoihop"]').forEach(cb => cb.checked = false);
    document.querySelectorAll('input[data-group="xemdebiet"]').forEach(cb => cb.checked = false);
    document.getElementById('cdYeuCauTraLoi').checked = false;
    document.getElementById('cdChuyenThuKy').checked = false;
    document.getElementById('cdTheoDoiVB').checked = false;
    document.querySelectorAll('#modalChiDao .pb-row.suggested').forEach(r => r.classList.remove('suggested'));
    
    // Auto-fill from AI suggestion if available
    if (selectedDoc && suggestedDocs[selectedDoc.id]) {
      const s = suggestedDocs[selectedDoc.id];
      s.chuTri.forEach(name => {
        const cb = document.querySelector(`#modalChiDao input[data-group="chutri"][value="${name}"]`);
        if (cb) { cb.checked = true; cb.closest('.pb-row')?.classList.add('suggested'); }
      });
      s.phoiHop.forEach(name => {
        const cb = document.querySelector(`#modalChiDao input[data-group="phoihop"][value="${name}"]`);
        if (cb) { cb.checked = true; cb.closest('.pb-row')?.classList.add('suggested'); }
      });
    }
  }
}

async function saveChiDao() {
  if (!selectedDoc) return;
  
  // Prevent double-click
  const saveBtn = document.getElementById('btnSaveChiDao');
  if (saveBtn.disabled) return;
  saveBtn.disabled = true;
  saveBtn.textContent = 'Đang lưu...';

  try {

  // Gather form data
  const noiDung = document.getElementById('cdNoiDung').value.trim();
  const hanGiaiQuyet = document.getElementById('cdHanGiaiQuyet').value;
  const doKhan = document.getElementById('cdDoKhan').value || 'Bình thường';
  const yeuCauTraLoi = document.getElementById('cdYeuCauTraLoi').checked;
  const chuyenThuKy = document.getElementById('cdChuyenThuKy').checked;
  const theoDoiVanBan = document.getElementById('cdTheoDoiVB').checked;

  // Get chủ trì (checkboxes)
  const chuTri = [];
  document.querySelectorAll('input[data-group="chutri"]:checked').forEach(cb => {
    chuTri.push(cb.value);
  });

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

  // Check if this document already has a chi dao (update mode)
  const existingChiDao = chiDaoData[selectedDoc.id] && chiDaoData[selectedDoc.id].length > 0;
  const existingId = existingChiDao ? chiDaoData[selectedDoc.id][0].id : '';

  const data = {
    action: existingChiDao ? 'updateChiDao' : 'saveChiDao',
    existingId: existingId,
    vanBanId: selectedDoc.id,
    noiDung,
    chuTri: chuTri.join(', '),
    phoiHop: phoiHop.join(', '),
    xemDeBiet: xemDeBiet.join(', '),
    hanGiaiQuyet,
    doKhan,
    yeuCauTraLoi,
    chuyenThuKy,
    theoDoiVanBan,
    nguoiChiDao: 'Nguyễn Ngọc Tuyến',
  };

  // Save to Google Sheets
  if (GAS_URL) {
    try {
      const resp = await fetch(GAS_URL, {
        method: 'POST',
        redirect: 'follow',
        body: JSON.stringify(data),
      });
      const text = await resp.text();
      let result;
      try { result = JSON.parse(text); } catch (pe) { result = { success: true }; }
      if (result.success || resp.ok) {
        showToast(existingChiDao ? 'Đã cập nhật chỉ đạo!' : 'Đã lưu chỉ đạo thành công!', 'success');
        // Update local cache
        if (existingChiDao) {
          chiDaoData[selectedDoc.id] = [data];
        } else {
          if (!chiDaoData[selectedDoc.id]) chiDaoData[selectedDoc.id] = [];
          chiDaoData[selectedDoc.id].push(data);
        }
        delete suggestedDocs[selectedDoc.id]; // Clear suggested state
        closeChiDaoModal();
        renderDocumentList(); // Update status dots
        if (selectedDoc) updateAssignmentSummary(selectedDoc); // Refresh panel
        return;
      } else {
        showToast('Lỗi: ' + (result.error || 'Không xác định'), 'error');
      }
    } catch (e) {
      console.error('Save chi dao error:', e);
      showToast('Lỗi kết nối server: ' + e.message, 'error');
    }
  } else {
    // Save locally (demo mode)
    if (!chiDaoData[selectedDoc.id]) chiDaoData[selectedDoc.id] = [];
    chiDaoData[selectedDoc.id].push(data);
    // Clear suggested state after saving
    delete suggestedDocs[selectedDoc.id];
    showToast('Đã lưu chỉ đạo (chế độ offline)', 'success');
    closeChiDaoModal();
    renderDocumentList();
    if (selectedDoc) updateAssignmentSummary(selectedDoc); // Refresh panel
  }

  } finally {
    // Re-enable save button
    saveBtn.disabled = false;
    saveBtn.innerHTML = '<span>Lưu</span>';
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
  document.querySelectorAll('#modalChuyenChiDao input[type="checkbox"]').forEach(cb => cb.checked = false);
}

function closeChuyenChiDaoModal() {
  document.getElementById('modalChuyenChiDao').classList.remove('show');
  document.body.style.overflow = '';
}

async function saveChuyenChiDao() {
  if (!selectedDoc) return;

  const noiDungChuyen = document.getElementById('ccdNoiDung').value.trim();
  
  const ccdChuTri = [];
  document.querySelectorAll('#modalChuyenChiDao input[data-group="ccd_chutri"]:checked').forEach(cb => ccdChuTri.push(cb.value));
  const chuTri = ccdChuTri.join(', ');

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
    nguoiChuyen: 'Nguyễn Ngọc Tuyến',
  };

  if (GAS_URL) {
    try {
      const resp = await fetch(GAS_URL, {
        method: 'POST',
        redirect: 'follow',
        body: JSON.stringify(data),
      });
      const text = await resp.text();
      let result;
      try { result = JSON.parse(text); } catch (pe) { result = { success: true }; }
      if (result.success || resp.ok) {
        showToast('Đã lưu chuyển chỉ đạo!', 'success');
        if (!chuyenChiDaoData[selectedDoc.id]) chuyenChiDaoData[selectedDoc.id] = [];
        chuyenChiDaoData[selectedDoc.id].push(data);
        closeChuyenChiDaoModal();
        return;
      } else {
        showToast('Lỗi: ' + (result.error || 'Không xác định'), 'error');
      }
    } catch (e) {
      showToast('Lỗi kết nối server: ' + e.message, 'error');
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
  // Ban giám đốc (dynamic)
  const bgdBody = document.getElementById('pb-bgd');
  ORG_STRUCTURE.banGiamDocCCD.forEach((name, i) => {
    bgdBody.appendChild(createPBRow(name, i % 2 === 1));
  });

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
      <span class="pb-check-cell">
        <input type="checkbox" data-group="chutri" value="${escapeHtml(name)}">
      </span>
      <span class="pb-check-cell">
        <input type="checkbox" data-group="phoihop" value="${escapeHtml(name)}">
      </span>
      <span class="pb-check-cell">
        <input type="checkbox" data-group="xemdebiet" value="${escapeHtml(name)}">
      </span>
    </div>
  `;
  // Mutual exclusion per row + only 1 chủ trì globally
  row.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) {
        // Uncheck other groups in same row (mutual exclusion per person)
        row.querySelectorAll('input[type="checkbox"]').forEach(other => {
          if (other !== cb) other.checked = false;
        });
        // If chutri selected, uncheck all other chutri globally (only 1 chủ trì)
        if (cb.dataset.group === 'chutri') {
          document.querySelectorAll('#modalChiDao input[data-group="chutri"]').forEach(other => {
            if (other !== cb) other.checked = false;
          });
        }
      }
    });
  });
  return row;
}

function createCCDRow(name, isAlt) {
  const row = document.createElement('div');
  row.className = 'ccd-row' + (isAlt ? ' alt' : '');
  row.innerHTML = `
    <span class="ccd-name">${escapeHtml(name)}</span>
    <span class="ccd-check-cell">
      <input type="checkbox" data-group="ccd_chutri" value="${escapeHtml(name)}">
    </span>
    <span class="ccd-check-cell">
      <input type="checkbox" data-group="ccd_xemdebiet" value="${escapeHtml(name)}">
    </span>
  `;
  // Mutual exclusion per row
  row.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.addEventListener('change', () => {
      if (cb.checked) {
        row.querySelectorAll('input[type="checkbox"]').forEach(other => {
          if (other !== cb) other.checked = false;
        });
      }
    });
  });
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

// ============================================================
// OVERVIEW MODAL
// ============================================================

function openOverviewModal() {
  const modal = document.getElementById('modalOverview');
  modal.style.display = 'flex';
  populateOverviewTable();
}

function closeOverviewModal() {
  document.getElementById('modalOverview').style.display = 'none';
}

function populateOverviewTable() {
  const tbody = document.getElementById('overviewTableBody');
  const emptyMsg = document.getElementById('overviewEmpty');
  const table = document.getElementById('overviewTable');
  tbody.innerHTML = '';

  // Find all documents that have chi dao
  const assignedDocs = allDocuments.filter(doc => 
    (chiDaoData[doc.id] && chiDaoData[doc.id].length > 0) || suggestedDocs[doc.id]
  );

  if (assignedDocs.length === 0) {
    table.style.display = 'none';
    emptyMsg.style.display = 'block';
    return;
  }

  table.style.display = 'table';
  emptyMsg.style.display = 'none';

  assignedDocs.forEach((doc, idx) => {
    const cdArr = chiDaoData[doc.id] || [];
    const cd = cdArr[0] || suggestedDocs[doc.id] || {}; // fallback to suggestion data
    
    // Get chuyen chi dao info
    const ccdList = chuyenChiDaoData[doc.id] || [];
    const chuyenCD = ccdList.map(c => c.nguoiChuyen || c.phoGiamDoc || '').filter(Boolean).join(', ');
    
    const isSuggested = suggestedDocs[doc.id] && !(chiDaoData[doc.id]?.length > 0);
    const tr = document.createElement('tr');
    tr.dataset.docId = doc.id;
    if (isSuggested) tr.classList.add('suggested-row');
    const displayChuTri = Array.isArray(cd.chuTri) ? cd.chuTri.join(', ') : (cd.chuTri || '—');
    const displayPhoiHop = Array.isArray(cd.phoiHop) ? cd.phoiHop.join(', ') : (cd.phoiHop || '—');
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><span class="overview-so-vb">${escapeHtml(doc.soVanBan)}</span></td>
      <td><div class="overview-trich-yeu">${escapeHtml(doc.trichYeu || 'Không có trích yếu')}</div></td>
      <td>${escapeHtml(chuyenCD || '—')}</td>
      <td>${escapeHtml(displayChuTri)}</td>
      <td>${escapeHtml(displayPhoiHop)}</td>
    `;
    
    tr.addEventListener('click', () => {
      closeOverviewModal();
      // Find and select the document
      const targetDoc = allDocuments.find(d => d.id === doc.id);
      if (targetDoc) {
        selectDocument(targetDoc);
        // Small delay to let document select, then open chi dao modal
        setTimeout(() => openChiDaoModal(), 200);
      }
    });
    
    tbody.appendChild(tr);
  });
}
