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

// Source patterns for document origin detection
const SOURCE_PATTERNS = {
  diaPhuong: ['ubnd', 'phường', 'xã', 'thị xã', 'thành phố vũng tàu', 'thành phố bà rịa', 'phú mỹ', 'long điền', 'đất đỏ', 'xuyên mộc', 'châu đức', 'côn đảo', 'phước thắng', 'tam thắng', 'thắng nhất', 'long hương'],
  dichVuDienLuc: ['dịch vụ điện lực', 'dvđl', 'dvdl'],
  trungTamDieuDo: ['điều độ', 'đđhtđ', 'trung tâm điều độ'],
  congAnPCCC: ['công an', 'cảnh sát', 'pccc', 'cnch'],
};

// Content keyword groups
const CONTENT_KEYWORDS = {
  cuongChe: ['cưỡng chế', 'thu hồi đất', 'giải phóng mặt bằng', 'phương án cưỡng chế'],
  dienMatTroi: ['năng lượng mặt trời', 'mặt trời mái nhà', 'nlmt', 'điện mặt trời', 'solar', 'tự sản xuất', 'tự tiêu thụ', 'đấu nối với hệ thống điện quốc gia'],
  giaoThong: ['giao thông', 'đồng bộ giao thông', 'đường', 'di dời', 'lê quang định', 'bình giã'],
  cungCapDien: ['cung cấp điện', 'cấp điện', 'đề nghị cấp điện', 'khôi phục điện', 'phối hợp cung cấp điện'],
  tuyenTruyen: ['tuyên truyền', 'tuyên truyền điện'],
  hotline: ['hotline', 'liveline', 'thi công đường dây', 'mang điện', 'phương án thi công', 'sự cố', 'xlsc'],
  vanHanh: ['vận hành', 'phương thức vận hành', 'hệ thống điện', 'cắt điện', 'sa thải', 'relay', 'f81', 'đánh số thiết bị', 'trạm biến thế'],
  phuongAnCCD: ['phương án đảm bảo', 'phương án cung cấp điện', 'cung cấp điện tháng', 'cung cấp điện phục vụ lễ', 'quốc khánh'],
  keHoach: ['kế hoạch cắt điện', 'kế hoạch bảo dưỡng', 'sửa chữa', 'bảo dưỡng'],
  congSuatDMTMN: ['công suất tối đa', 'đmtmn', 'phân bổ công suất'],
  dauTuXD: ['đầu tư', 'xây dựng', 'tái định cư', 'hạ tầng kỹ thuật', 'gói thầu', 'khối lượng vật tư thu hồi'],
  phatTrienLuoi: ['phát triển lưới', 'tăng cường công suất', 'trạm biến áp', 'nâng cấp lưới', 'mở rộng lưới'],
  conDao: ['côn đảo', 'đặc khu côn đảo'],
  keHoachVatTu: ['kế hoạch', 'vật tư', 'vận chuyển', 'phụ tùng'],
  kinhDoanh: ['kinh doanh', 'khách hàng', 'giá điện', 'tiền điện', 'mua bán điện', 'nghị định 243', 'nd-cp'],
  keToan: ['tài chính', 'kế toán', 'chi phí', 'ngân sách', 'phúc lợi'],
  nhanSu: ['tuyển dụng', 'nhân sự', 'vị trí chức danh', 'khung năng lực', 'cơ cấu tổ chức', 'nâng bậc', 'giữ bậc', 'đào tạo', 'thạc sĩ', 'tuyển sinh', 'sát hạch'],
  dangUy: ['đảng', 'nghị quyết', 'bch', 'trung ương', 'đảng ủy', 'chi bộ', 'quán triệt'],
  anToan: ['an toàn', 'an toàn lao động', 'tai nạn lao động', 'phòng thủ dân sự'],
  phapLy: ['luật', 'pháp luật', 'an ninh dữ liệu', 'dự án luật', 'góp ý'],
  thietTri: ['thiết trí', 'đấu nối nhánh dây', 'mắc điện khách hàng', 'chuẩn hoá'],
  giayMoi: ['giấy mời', 'hội nghị', 'họp'],
};

function matchKeywords(text, kwGroup) {
  return kwGroup.some(kw => text.includes(kw.toLowerCase()));
}

function detectSource(text) {
  for (const [src, patterns] of Object.entries(SOURCE_PATTERNS)) {
    if (patterns.some(p => text.includes(p.toLowerCase()))) return src;
  }
  return 'other';
}

function suggestAssignment(doc) {
  const text = ((doc.trichYeu || '') + ' ' + (doc.coQuanBanHanh || '')).toLowerCase();
  const source = detectSource(text);

  // === 1. VB địa phương ===
  if (source === 'diaPhuong') {
    if (matchKeywords(text, CONTENT_KEYWORDS.cuongChe)) {
      return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KD - Đặng Quang Trung'], ccdXemDeBiet: [] };
    }
    if (matchKeywords(text, CONTENT_KEYWORDS.dienMatTroi)) {
      return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KD - Đặng Quang Trung'], ccdXemDeBiet: ['PGĐ KT - Trần Thanh Hải'] };
    }
    if (matchKeywords(text, CONTENT_KEYWORDS.giaoThong)) {
      const hasLargeScope = matchKeywords(text, ['gói thầu', 'khối lượng', 'giai đoạn', 'dự án', 'hạ tầng']);
      if (hasLargeScope) {
        return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ ĐTXD - Lại Xuân Phương'], ccdXemDeBiet: ['PGĐ KT - Trần Thanh Hải'] };
      } else {
        return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: ['PGĐ ĐTXD - Lại Xuân Phương'] };
      }
    }
    if (matchKeywords(text, CONTENT_KEYWORDS.cungCapDien)) {
      return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KD - Đặng Quang Trung'], ccdXemDeBiet: ['PGĐ KT - Trần Thanh Hải'] };
    }
    if (matchKeywords(text, CONTENT_KEYWORDS.tuyenTruyen)) {
      return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: ['PGĐ KD - Đặng Quang Trung'] };
    }
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: [] };
  }

  // === 2. VB Dịch vụ Điện lực (Hotline) ===
  if (source === 'dichVuDienLuc') {
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: [] };
  }

  // === 3. VB Trung tâm Điều độ ===
  if (source === 'trungTamDieuDo') {
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: [] };
  }

  // === 4. VB Công an/PCCC ===
  if (source === 'congAnPCCC') {
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: [] };
  }

  // === 5. Đầu tư xây dựng ===
  if (matchKeywords(text, CONTENT_KEYWORDS.dauTuXD)) {
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ ĐTXD - Lại Xuân Phương'], ccdXemDeBiet: [] };
  }

  // === 6. Phát triển lưới, trạm điện → Chỉ đạo thẳng P.KTAT ===
  if (matchKeywords(text, CONTENT_KEYWORDS.phatTrienLuoi)) {
    return { type: 'chiDaoThang', chuTri: ['Phòng Kỹ thuật & An Toàn'], phoiHop: [], xemDeBiet: [] };
  }

  // === 7. Côn Đảo → Chỉ đạo thẳng ===
  if (matchKeywords(text, CONTENT_KEYWORDS.conDao)) {
    return { type: 'chiDaoThang', chuTri: ['Điện lực Đặc khu Côn Đảo'], phoiHop: [], xemDeBiet: [] };
  }

  // === 8. Kế hoạch & Vật tư → Chỉ đạo thẳng ===
  if (matchKeywords(text, CONTENT_KEYWORDS.keHoachVatTu)) {
    return { type: 'chiDaoThang', chuTri: ['Phòng Kế hoạch & Vật Tư'], phoiHop: [], xemDeBiet: [] };
  }

  // === 9. Kinh doanh → Chỉ đạo thẳng ===
  if (matchKeywords(text, CONTENT_KEYWORDS.kinhDoanh)) {
    return { type: 'chiDaoThang', chuTri: ['Phòng Kinh doanh'], phoiHop: [], xemDeBiet: [] };
  }

  // === 10. Kế toán → Chỉ đạo thẳng ===
  if (matchKeywords(text, CONTENT_KEYWORDS.keToan)) {
    return { type: 'chiDaoThang', chuTri: ['Phòng Tài chính Kế toán'], phoiHop: [], xemDeBiet: [] };
  }

  // === 11. Đảng ủy → Chỉ đạo thẳng ===
  if (matchKeywords(text, CONTENT_KEYWORDS.dangUy)) {
    return { type: 'chiDaoThang', chuTri: ['BCH Đảng ủy'], phoiHop: [], xemDeBiet: [] };
  }

  // === 12. Điện mặt trời (generic) ===
  if (matchKeywords(text, CONTENT_KEYWORDS.dienMatTroi)) {
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KD - Đặng Quang Trung'], ccdXemDeBiet: ['PGĐ KT - Trần Thanh Hải'] };
  }

  // === 13. Hotline/Liveline (generic) ===
  if (matchKeywords(text, CONTENT_KEYWORDS.hotline)) {
    return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: [] };
  }

  // === DEFAULT: VB khác, giấy mời → Chuyển CĐ PGĐ KT ===
  return { type: 'chuyenChiDao', ccdChuTri: ['PGĐ KT - Trần Thanh Hải'], ccdXemDeBiet: [] };
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
  
  // If chi dao modal is open with a selected doc, apply checkboxes
  if (selectedDoc && suggestedDocs[selectedDoc.id]) {
    const s = suggestedDocs[selectedDoc.id];
    if (s.type === 'chiDaoThang') {
      applySuggestionToChiDaoModal(s);
    }
  }
}

function applySuggestionToChiDaoModal(s) {
  document.querySelectorAll('#modalChiDao input[data-group="chutri"]').forEach(cb => cb.checked = false);
  document.querySelectorAll('#modalChiDao input[data-group="phoihop"]').forEach(cb => cb.checked = false);
  document.querySelectorAll('#modalChiDao input[data-group="xemdebiet"]').forEach(cb => cb.checked = false);
  
  (s.chuTri || []).forEach(name => {
    const cb = document.querySelector(`#modalChiDao input[data-group="chutri"][value="${name}"]`);
    if (cb) { cb.checked = true; cb.closest('.pb-row')?.classList.add('suggested'); }
  });
  (s.phoiHop || []).forEach(name => {
    const cb = document.querySelector(`#modalChiDao input[data-group="phoihop"][value="${name}"]`);
    if (cb) { cb.checked = true; cb.closest('.pb-row')?.classList.add('suggested'); }
  });
}

function applySuggestionToCCDModal(s) {
  document.querySelectorAll('#modalChuyenChiDao input[type="checkbox"]').forEach(cb => cb.checked = false);
  
  (s.ccdChuTri || []).forEach(name => {
    const cb = document.querySelector(`#modalChuyenChiDao input[data-group="ccd_chutri"][value="${name}"]`);
    if (cb) { cb.checked = true; cb.closest('.ccd-row')?.classList.add('suggested'); }
  });
  (s.ccdXemDeBiet || []).forEach(name => {
    const cb = document.querySelector(`#modalChuyenChiDao input[data-group="ccd_xemdebiet"][value="${name}"]`);
    if (cb) { cb.checked = true; cb.closest('.ccd-row')?.classList.add('suggested'); }
  });
}

// ============================================================
// INITIALIZATION
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
  initUI();
  loadDocuments();
  loadDriveMapping();
  populateDynamicForms();
  initXDB();
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
  document.getElementById('tabChuyenCDView').addEventListener('click', () => { switchModalTab('chuyencd'); loadTab2CCD(); });
  document.getElementById('tabThongTinVB').addEventListener('click', () => switchModalTab('thongtin'));

  document.getElementById('btnSaveChiDao').addEventListener('click', saveChiDao);
  document.getElementById('btnSaveChuyenChiDao').addEventListener('click', saveChuyenChiDao);
  document.getElementById('btnSaveTab2CCD').addEventListener('click', saveTab2CCD);
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
    // Always load documents from documents.json (source of truth)
    await loadFromJSON();
    
    // Load chi dao data from GAS (if available)
    if (GAS_URL) {
      loadChiDaoFromSheets();
      loadDriveMapping();
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
    ${doc.ngayTaiLen ? `<div class="doc-card-upload-time"><span class="material-icons-outlined" style="font-size:12px">cloud_upload</span> ${formatUploadTime(doc.ngayTaiLen)}</div>` : ''}
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
  
  if (!cd && !suggested) {
    panel.style.display = 'none';
    return;
  }
  
  panel.style.display = 'block';
  const isSuggested = !cd && suggested;
  let html = '';
  
  if (cd) {
    // Show saved chi dao data
    const chuTri = cd.chuTri || '';
    if (chuTri) html += `<div class="assign-item"><span class="assign-label ct">Chủ trì:</span> <span class="assign-value">${escapeHtml(chuTri)}</span></div>`;
    const phoiHop = cd.phoiHop || '';
    if (phoiHop) html += `<div class="assign-item"><span class="assign-label ph">Phối hợp:</span> <span class="assign-value">${escapeHtml(phoiHop)}</span></div>`;
    const xdb = cd.xemDeBiet || '';
    if (xdb) html += `<div class="assign-item"><span class="assign-label xdb">Xem để biết:</span> <span class="assign-value">${escapeHtml(xdb)}</span></div>`;
    html = '<div class="assign-badge saved">✅ Đã phân công</div>' + html;
  } else if (suggested) {
    if (suggested.type === 'chuyenChiDao') {
      const ccdCT = (suggested.ccdChuTri || []).join(', ');
      if (ccdCT) html += `<div class="assign-item"><span class="assign-label ct">Chuyển CĐ:</span> <span class="assign-value">${escapeHtml(ccdCT)}</span></div>`;
      const ccdXDB = (suggested.ccdXemDeBiet || []).join(', ');
      if (ccdXDB) html += `<div class="assign-item"><span class="assign-label ph">Phối hợp:</span> <span class="assign-value">${escapeHtml(ccdXDB)}</span></div>`;
    } else {
      const chuTri = (suggested.chuTri || []).join(', ');
      if (chuTri) html += `<div class="assign-item"><span class="assign-label ct">CĐ thẳng:</span> <span class="assign-value">${escapeHtml(chuTri)}</span></div>`;
    }
    html = '<div class="assign-badge suggested">🤖 Đề xuất AI</div>' + html;
  }
  
  if (!html) {
    panel.style.display = 'none';
    return;
  }
  
  body.innerHTML = html;
}

// ============================================================
// FILTER & SEARCH
// ============================================================

function filterDocuments(query) {
  // Search + tab + filter mode all handled by applyFilterMode
  applyFilterMode();
}

let currentFilterMode = 'all'; // 'all', 'assigned', 'unassigned'

let currentTab = 'chutri';

function switchTab(tab) {
  currentTab = tab;
  document.getElementById('tabChuTri').classList.toggle('active', tab === 'chutri');
  document.getElementById('tabPhoiHop').classList.toggle('active', tab === 'phoihop');
  applyFilterMode();
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
  
  // Apply tab filter (Chủ trì / Phối hợp)
  if (currentTab === 'chutri') {
    docs = docs.filter(d => !d.loaiVB || d.loaiVB === 'Chủ trì');
  } else if (currentTab === 'phoihop') {
    docs = docs.filter(d => d.loaiVB === 'Phối hợp');
  }
  
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
  const chuTriDocs = allDocuments.filter(d => !d.loaiVB || d.loaiVB === 'Chủ trì');
  const phoiHopDocs = allDocuments.filter(d => d.loaiVB === 'Phối hợp');
  document.getElementById('countChuTri').textContent = chuTriDocs.length;
  document.getElementById('countPhoiHop').textContent = phoiHopDocs.length;
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
  document.getElementById('tabChuyenCDView').classList.toggle('active', tab === 'chuyencd');
  document.getElementById('tabThongTinVB').classList.toggle('active', tab === 'thongtin');
  document.getElementById('contentChiDao').classList.toggle('active', tab === 'chidao');
  document.getElementById('contentChuyenCD').classList.toggle('active', tab === 'chuyencd');
  document.getElementById('contentThongTin').classList.toggle('active', tab === 'thongtin');
}

function loadTab2CCD() {
  if (!selectedDoc) return;
  
  // Reset form
  document.getElementById('tab2CcdNoiDung').value = '';
  document.querySelectorAll('#contentChuyenCD input[type="checkbox"]').forEach(cb => cb.checked = false);
  
  // Load existing CCD data
  const ccdList = chuyenChiDaoData[selectedDoc.id] || [];
  if (ccdList.length > 0) {
    const latest = ccdList[ccdList.length - 1];
    document.getElementById('tab2CcdNoiDung').value = latest.noiDungChuyen || '';
    
    if (latest.chuTri) {
      const ctList = latest.chuTri.split(',').map(s => s.trim());
      document.querySelectorAll('#contentChuyenCD input[data-group="tab2_ccd_chutri"]').forEach(cb => {
        cb.checked = ctList.includes(cb.value);
      });
    }
    if (latest.xemDeBiet) {
      const xdbList = latest.xemDeBiet.split(',').map(s => s.trim());
      document.querySelectorAll('#contentChuyenCD input[data-group="tab2_ccd_xemdebiet"]').forEach(cb => {
        cb.checked = xdbList.includes(cb.value);
      });
    }
  } else {
    // Auto-fill from AI suggestion
    const suggest = suggestedDocs[selectedDoc.id];
    if (suggest && suggest.type === 'chuyenChiDao') {
      (suggest.ccdChuTri || []).forEach(name => {
        const cb = document.querySelector(`#contentChuyenCD input[data-group="tab2_ccd_chutri"][value="${name}"]`);
        if (cb) { cb.checked = true; cb.closest('.ccd-row')?.classList.add('suggested'); }
      });
      (suggest.ccdXemDeBiet || []).forEach(name => {
        const cb = document.querySelector(`#contentChuyenCD input[data-group="tab2_ccd_xemdebiet"][value="${name}"]`);
        if (cb) { cb.checked = true; cb.closest('.ccd-row')?.classList.add('suggested'); }
      });
    }
  }
}

async function saveTab2CCD() {
  if (!selectedDoc) return;
  
  const saveBtn = document.getElementById('btnSaveTab2CCD');
  if (saveBtn.disabled) return;
  saveBtn.disabled = true;
  saveBtn.textContent = 'Đang lưu...';

  try {
    const noiDungChuyen = document.getElementById('tab2CcdNoiDung').value.trim();
    
    const ccdChuTri = [];
    document.querySelectorAll('#contentChuyenCD input[data-group="tab2_ccd_chutri"]:checked').forEach(cb => ccdChuTri.push(cb.value));
    const chuTri = ccdChuTri.join(', ');

    const xemDeBiet = [];
    document.querySelectorAll('#contentChuyenCD input[data-group="tab2_ccd_xemdebiet"]:checked').forEach(cb => xemDeBiet.push(cb.value));

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
        await fetch(GAS_URL, { method: 'POST', mode: 'no-cors', body: JSON.stringify(data) });
      } catch (e) { console.warn('Save CCD error:', e); }
    }

    // Update local data
    if (!chuyenChiDaoData[selectedDoc.id]) chuyenChiDaoData[selectedDoc.id] = [];
    const existing = chuyenChiDaoData[selectedDoc.id];
    if (existing.length > 0) {
      existing[existing.length - 1] = { ...data, timestamp: new Date().toISOString() };
    } else {
      existing.push({ ...data, timestamp: new Date().toISOString() });
    }

    // Clear AI suggestion for this doc
    delete suggestedDocs[selectedDoc.id];
    document.querySelectorAll('#contentChuyenCD .ccd-row.suggested').forEach(r => r.classList.remove('suggested'));

    showToast('✅ Đã lưu chuyển chỉ đạo', 'success');
    renderDocumentList();
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = 'Lưu chuyển chỉ đạo';
  }
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
    
    // Auto-fill from AI suggestion if available (chiDaoThang only)
    if (selectedDoc && suggestedDocs[selectedDoc.id]) {
      const s = suggestedDocs[selectedDoc.id];
      if (s.type === 'chiDaoThang') {
        applySuggestionToChiDaoModal(s);
      }
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
  document.querySelectorAll('#modalChuyenChiDao .ccd-row.suggested').forEach(r => r.classList.remove('suggested'));

  // Auto-fill from AI suggestion if available (chuyenChiDao only)
  if (selectedDoc && suggestedDocs[selectedDoc.id]) {
    const s = suggestedDocs[selectedDoc.id];
    if (s.type === 'chuyenChiDao') {
      applySuggestionToCCDModal(s);
    }
  }
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

  // Tab2 CCD - Ban giám đốc (inside chi dao modal)
  const tab2BGD = document.getElementById('tab2-ccd-bgd-body');
  ORG_STRUCTURE.banGiamDocCCD.forEach((name, i) => {
    tab2BGD.appendChild(createCCDRow(name, i % 2 === 1, 'tab2_ccd'));
  });

  // Tab2 CCD - Đảng ủy & Đoàn thể
  const tab2DUDT = document.getElementById('tab2-ccd-dudt-body');
  ORG_STRUCTURE.dangUyDoanThe.forEach((name, i) => {
    tab2DUDT.appendChild(createCCDRow(name, i % 2 === 1, 'tab2_ccd'));
  });

  // Wire up group header toggle checkboxes
  initGroupToggle();
}

function initGroupToggle() {
  document.querySelectorAll('input[data-group-toggle]').forEach(headerCb => {
    headerCb.addEventListener('change', () => {
      const groupName = headerCb.dataset.groupToggle; // chutri, phoihop, xemdebiet
      const targetId = headerCb.dataset.target; // pb-bgd, pb-pbdv, etc.
      const targetBody = document.getElementById(targetId);
      if (!targetBody) return;

      const isChecked = headerCb.checked;

      if (groupName === 'chutri' && isChecked) {
        // Chutri: only 1 globally → uncheck all other chutri first
        document.querySelectorAll('#modalChiDao input[data-group="chutri"]').forEach(cb => {
          cb.checked = false;
        });
        // Also uncheck other group toggle chutri headers
        document.querySelectorAll('input[data-group-toggle="chutri"]').forEach(other => {
          if (other !== headerCb) other.checked = false;
        });
        // For chutri, only check the FIRST child (can't have multiple chutri)
        const firstChuTri = targetBody.querySelector('input[data-group="chutri"]');
        if (firstChuTri) {
          firstChuTri.checked = true;
          // Uncheck other groups in that row
          const row = firstChuTri.closest('.pb-row');
          if (row) {
            row.querySelectorAll('input[type="checkbox"]').forEach(other => {
              if (other !== firstChuTri) other.checked = false;
            });
          }
        }
      } else {
        // Phoihop or Xemdebiet: toggle all children
        targetBody.querySelectorAll(`input[data-group="${groupName}"]`).forEach(cb => {
          cb.checked = isChecked;
          if (isChecked) {
            // Uncheck other groups in same row (mutual exclusion per row)
            const row = cb.closest('.pb-row');
            if (row) {
              row.querySelectorAll('input[type="checkbox"]').forEach(other => {
                if (other !== cb) other.checked = false;
              });
            }
          }
        });
        // Also uncheck other header toggles in same group header (mutual exclusion)
        if (isChecked) {
          const headerDiv = headerCb.closest('.pb-cols');
          if (headerDiv) {
            headerDiv.querySelectorAll('input[data-group-toggle]').forEach(other => {
              if (other !== headerCb) other.checked = false;
            });
          }
        }
      }
    });

    // Prevent header checkbox click from toggling the group body collapse
    headerCb.addEventListener('click', (e) => {
      e.stopPropagation();
    });
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

function createCCDRow(name, isAlt, prefix = 'ccd') {
  const row = document.createElement('div');
  row.className = 'ccd-row' + (isAlt ? ' alt' : '');
  row.innerHTML = `
    <span class="ccd-name">${escapeHtml(name)}</span>
    <span class="ccd-check-cell">
      <input type="checkbox" data-group="${prefix}_chutri" value="${escapeHtml(name)}">
    </span>
    <span class="ccd-check-cell">
      <input type="checkbox" data-group="${prefix}_xemdebiet" value="${escapeHtml(name)}">
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

function formatUploadTime(val) {
  if (!val) return '';
  // Already dd/MM/yyyy HH:mm format
  if (/^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}$/.test(val)) return val;
  // ISO format → parse and format with VN timezone
  try {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      // Adjust to UTC+7
      const vn = new Date(d.getTime() + 7 * 60 * 60 * 1000);
      const dd = String(vn.getUTCDate()).padStart(2, '0');
      const mm = String(vn.getUTCMonth() + 1).padStart(2, '0');
      const yyyy = vn.getUTCFullYear();
      const hh = String(vn.getUTCHours()).padStart(2, '0');
      const mi = String(vn.getUTCMinutes()).padStart(2, '0');
      return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
    }
  } catch (e) {}
  return String(val);
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
    const cd = cdArr[0] || null;
    const suggest = suggestedDocs[doc.id] || null;
    
    // Get saved chuyen chi dao info
    const ccdList = chuyenChiDaoData[doc.id] || [];
    const savedChuyenCD = ccdList.map(c => c.chuTri || '').filter(Boolean).join(', ');
    
    const isSuggested = suggest && !cd;
    const tr = document.createElement('tr');
    tr.dataset.docId = doc.id;
    if (isSuggested) tr.classList.add('suggested-row');

    let displayChuyenCD = savedChuyenCD || '—';
    let displayChuTri = '—';
    let displayPhoiHop = '—';

    if (cd) {
      // Saved chi dao data
      displayChuTri = cd.chuTri || '—';
      displayPhoiHop = cd.phoiHop || '—';
    } else if (suggest) {
      if (suggest.type === 'chuyenChiDao') {
        // AI suggests chuyển chỉ đạo → show in Chuyển CĐ column
        displayChuyenCD = (suggest.ccdChuTri || []).join(', ') || '—';
        displayPhoiHop = (suggest.ccdXemDeBiet || []).join(', ') || '—';
        displayChuTri = '—';
      } else {
        // AI suggests chỉ đạo thẳng → show in Chủ trì column
        displayChuTri = (suggest.chuTri || []).join(', ') || '—';
        displayPhoiHop = (suggest.phoiHop || []).join(', ') || '—';
      }
    }

    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><span class="overview-so-vb">${escapeHtml(doc.soVanBan)}</span></td>
      <td><div class="overview-trich-yeu">${escapeHtml(doc.trichYeu || 'Không có trích yếu')}</div></td>
      <td>${escapeHtml(displayChuyenCD)}</td>
      <td>${escapeHtml(displayChuTri)}</td>
      <td>${escapeHtml(displayPhoiHop)}</td>
    `;
    
    tr.addEventListener('click', () => {
      closeOverviewModal();
      const targetDoc = allDocuments.find(d => d.id === doc.id);
      if (targetDoc) {
        selectDocument(targetDoc);
        setTimeout(() => openChiDaoModal(), 200);
      }
    });
    
    tbody.appendChild(tr);
  });
}

// ============================================================
// XEM ĐỂ BIẾT (XDB) MODULE
// ============================================================

const XDB_JSON_PATH = 'xdb_documents.json';
const XDB_FILES_BASE = '../DO-GĐ/SOURCE-DOGĐ/Xem để biết/';
let xdbDocuments = [];
let selectedXDBDoc = null;

async function loadXDBDocuments() {
  try {
    const resp = await fetch(XDB_JSON_PATH);
    if (!resp.ok) throw new Error('Failed to load XDB documents');
    xdbDocuments = await resp.json();
    renderXDBList();
    document.getElementById('xdbListLoading').style.display = 'none';
    // Update tab count
    document.getElementById('xdbCount').textContent = `(${xdbDocuments.length})`;
  } catch (e) {
    console.error('Error loading XDB docs:', e);
    document.getElementById('xdbListLoading').innerHTML = '<p style="color:var(--gray-400);text-align:center">Không tải được dữ liệu XDB</p>';
  }
}

function renderXDBList() {
  const container = document.getElementById('xdbList');
  // Remove old cards
  container.querySelectorAll('.doc-card').forEach(el => el.remove());
  
  xdbDocuments.forEach(doc => {
    const card = document.createElement('div');
    card.className = 'doc-card';
    card.dataset.id = doc.id;
    if (selectedXDBDoc && selectedXDBDoc.id === doc.id) card.classList.add('active');
    
    const trichYeu = doc.trichYeu || 'Không có trích yếu';
    // Clean trichYeu: remove inline chi dao text
    const cleanTY = trichYeu.replace(/\s+(Trần Thanh Hải|Đặng Quang Trung|Lại Xuân Phương|Nguyễn Ngọc Tuyến)\s*-\s*\d{2}\/\d{2}\/\d{4}.*$/i, '');
    
    card.innerHTML = `
      <div class="doc-card-header">
        <div class="doc-card-flag">
          <span class="doc-card-so">${escapeHtml(doc.soKyHieu)}</span>
        </div>
        <div class="doc-card-date-status">
          <span class="doc-card-date">${doc.ngayDen || ''}</span>
        </div>
      </div>
      <div class="doc-card-org">
        <span>${escapeHtml(doc.coQuanBanHanh || '—')}</span>
        <span class="doc-card-org-date">${doc.ngayVanBan || ''}</span>
      </div>
      <div class="doc-card-summary">${escapeHtml(cleanTY)}</div>
    `;
    
    card.addEventListener('click', () => selectXDBDocument(doc));
    container.appendChild(card);
  });
  
  document.getElementById('xdbListLoading').style.display = 'none';
}

function selectXDBDocument(doc) {
  selectedXDBDoc = doc;
  
  // Update list active state
  document.querySelectorAll('#xdbList .doc-card').forEach(el => el.classList.remove('active'));
  const activeCard = document.querySelector(`#xdbList .doc-card[data-id="${doc.id}"]`);
  if (activeCard) activeCard.classList.add('active');
  
  // Show preview
  document.getElementById('xdbPreviewPlaceholder').style.display = 'none';
  const previewContent = document.getElementById('xdbPreviewContent');
  previewContent.style.display = 'flex';
  previewContent.style.flexDirection = 'column';
  previewContent.style.flex = '1';
  previewContent.style.overflow = 'hidden';
  
  // Title & meta
  const cleanTY = (doc.trichYeu || '').replace(/\s+(Trần Thanh Hải|Đặng Quang Trung|Lại Xuân Phương|Nguyễn Ngọc Tuyến)\s*-\s*\d{2}\/\d{2}\/\d{4}.*$/i, '');
  document.getElementById('xdbPreviewTitle').textContent = doc.soKyHieu;
  document.getElementById('xdbPreviewMeta').innerHTML = `
    <div>${escapeHtml(doc.coQuanBanHanh)} · ${doc.ngayVanBan}</div>
    <div style="color:var(--gray-600);margin-top:4px">${escapeHtml(cleanTY)}</div>
  `;
  
  // File bar (like DOffice: "File văn bản: filename.pdf")
  const fileList = document.getElementById('xdbFileList');
  fileList.innerHTML = '';
  if (doc.files && doc.files.length > 0) {
    doc.files.forEach(file => {
      const pdfUrl = getPdfViewUrl('Xem để biết', file);
      const link = document.createElement('a');
      link.className = 'file-link';
      link.href = pdfUrl;
      link.target = '_blank';
      link.innerHTML = `<span class="material-icons-outlined" style="font-size:14px">picture_as_pdf</span> ${escapeHtml(file)}`;
      link.addEventListener('click', (e) => {
        e.preventDefault();
        document.getElementById('xdbFileViewer').innerHTML = `<iframe src="${pdfUrl}" title="PDF" style="width:100%;height:100%;border:none"></iframe>`;
      });
      fileList.appendChild(link);
    });
    
    // Auto-load first PDF
    const pdfFile = doc.files.find(f => f.toLowerCase().endsWith('.pdf'));
    if (pdfFile) {
      const autoUrl = getPdfViewUrl('Xem để biết', pdfFile);
      document.getElementById('xdbFileViewer').innerHTML = 
        `<iframe src="${autoUrl}" title="PDF" style="width:100%;height:100%;border:none"></iframe>`;
    }
  }
}

function openQuaTrinhXLModal() {
  if (!selectedXDBDoc) {
    showToast('Vui lòng chọn một văn bản', 'error');
    return;
  }
  
  const modal = document.getElementById('modalQuaTrinhXL');
  modal.style.display = 'flex';
  
  renderChiDaoChain(selectedXDBDoc.chiDaoChain || []);
}

function closeQuaTrinhXLModal() {
  document.getElementById('modalQuaTrinhXL').style.display = 'none';
}

function renderChiDaoChain(chain) {
  const body = document.getElementById('qtxlBody');
  
  if (chain.length === 0) {
    body.innerHTML = '<p style="text-align:center;color:var(--gray-400);padding:40px">Chưa có thông tin chỉ đạo</p>';
    return;
  }
  
  let html = '';
  chain.forEach(step => {
    const initials = getInitials(step.nguoiChiDao);
    const bgColor = getAvatarColor(step.nguoiChiDao);
    
    html += `<div class="chain-card">
      <div class="chain-card-header">
        <div class="chain-person">
          <div class="chain-avatar" style="background:${bgColor}">${initials}</div>
          <div>
            <div class="chain-name">${escapeHtml(step.nguoiChiDao)}</div>
            ${step.timestamp ? `<div class="chain-timestamp">${escapeHtml(step.nguoiChiDao.split(' - ')[0] || '')} - ${escapeHtml(step.timestamp)}</div>` : ''}
          </div>
        </div>
        <div class="chain-date">${escapeHtml(step.ngayGiao || '')}</div>
      </div>`;
    
    if (step.noiDung) {
      html += `<div class="chain-noidung">${escapeHtml(step.noiDung)}</div>`;
    }
    
    html += `<div class="chain-fields">`;
    html += `<div class="chain-field">
      <span class="chain-field-icon material-icons-outlined" style="font-size:16px">arrow_circle_right</span>
      <span class="chain-field-label">Chủ trì</span>
      <span class="chain-field-value">${escapeHtml(step.chuTri || '')}</span>
    </div>`;
    html += `<div class="chain-field">
      <span class="chain-field-icon material-icons-outlined" style="font-size:16px">group</span>
      <span class="chain-field-label">Phối hợp</span>
      <span class="chain-field-value">${escapeHtml(step.phoiHop || '')}</span>
    </div>`;
    html += `<div class="chain-field">
      <span class="chain-field-icon material-icons-outlined" style="font-size:16px">visibility</span>
      <span class="chain-field-label">Xem để biết</span>
      <span class="chain-field-value">${escapeHtml(step.xemDeBiet || '')}</span>
    </div>`;
    html += `<div class="chain-field">
      <span class="chain-field-icon material-icons-outlined" style="font-size:16px">event</span>
      <span class="chain-field-label">Hạn xử lý</span>
      <span class="chain-field-value deadline">${escapeHtml(step.hanXuLy || '')}</span>
    </div>`;
    html += `</div></div>`;
  });
  
  body.innerHTML = html;
}

function getInitials(name) {
  if (!name) return '?';
  const parts = name.split(/[\s-]+/).filter(Boolean);
  if (parts.length >= 2) return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  return name.substring(0, 2).toUpperCase();
}

function getAvatarColor(name) {
  const colors = ['#1565C0','#C62828','#2E7D32','#E65100','#4527A0','#00838F','#AD1457','#37474F'];
  let hash = 0;
  for (let i = 0; i < (name || '').length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
  return colors[Math.abs(hash) % colors.length];
}

// Sub-tab switching
function switchSubTab(tab) {
  // Update tab buttons
  document.getElementById('subTabChoXuLy').classList.toggle('active', tab === 'choxuly');
  document.getElementById('subTabXDB').classList.toggle('active', tab === 'xemdebiet');
  
  // Toggle Chờ xử lý elements
  const toolbar = document.getElementById('toolbarChoXuLy');
  const docList = document.getElementById('docListPanel');
  const docPreview = document.getElementById('docPreviewPanel');
  const toolbarActions = document.getElementById('toolbarActions');
  
  if (tab === 'choxuly') {
    toolbar.style.display = '';
    docList.style.display = '';
    docPreview.style.display = '';
  } else {
    toolbar.style.display = 'none';
    docList.style.display = 'none';
    docPreview.style.display = 'none';
    if (toolbarActions) toolbarActions.style.display = 'none';
  }
  
  // Toggle XDB section
  document.getElementById('sectionXDB').style.display = tab === 'xemdebiet' ? '' : 'none';
  
  // Load XDB data if not loaded
  if (tab === 'xemdebiet' && xdbDocuments.length === 0) {
    loadXDBDocuments();
  }
}

function initXDB() {
  // Sub-tab events
  document.getElementById('subTabChoXuLy').addEventListener('click', () => switchSubTab('choxuly'));
  document.getElementById('subTabXDB').addEventListener('click', () => switchSubTab('xemdebiet'));
  
  // QTXL modal events
  document.getElementById('btnQuaTrinhXL').addEventListener('click', openQuaTrinhXLModal);
  document.getElementById('closeQuaTrinhXL').addEventListener('click', closeQuaTrinhXLModal);
  document.getElementById('modalQuaTrinhXL').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) closeQuaTrinhXLModal();
  });
  
  // QTXL tab switching (dummy for now - only "Chỉ đạo và giao việc" has content)
  document.querySelectorAll('.qtxl-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.qtxl-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
    });
  });
  
  // Preload XDB count (fetch count without rendering)
  fetch(XDB_JSON_PATH).then(r => r.json()).then(data => {
    document.getElementById('xdbCount').textContent = `(${data.length})`;
  }).catch(() => {});
}
