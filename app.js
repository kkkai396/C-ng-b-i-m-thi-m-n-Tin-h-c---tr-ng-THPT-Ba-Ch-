/**
 * ============================================================================
 * CỔNG TRA CỨU & CÔNG BỐ ĐIỂM THI MÔN TIN HỌC - TRƯỜNG THPT BA CHẼ
 * File: app.js
 * Mô tả: Quản lý dữ liệu chuẩn biểu mẫu Bộ GD&ĐT (VNEDU/SMAS), xác thực Mã PIN,
 *        hiển thị bảng điểm 15 cột, thống kê, quản trị Admin & Import/Export Excel.
 * ============================================================================
 */

const STORAGE_KEY = 'THPT_BACHE_TINHOC_PORTAL_V2';
const LEGACY_STORAGE_KEY = 'THPT_BACHE_TINHOC_PORTAL_V1';
const VALID_SEMESTERS = ['Giữa kỳ I', 'Cuối kỳ I', 'Giữa kỳ II', 'Cuối kỳ II'];

/**
 * DANH SÁCH 43 HỌC SINH MẪU THEO ĐÚNG FILE EXCEL CỦA THẦY VI ANH QUÂN
 * Cấu trúc mỗi học sinh gồm 15 cột chuẩn biểu mẫu:
 * stt, mdd (Mã định danh Bộ GD&ĐT), name (Họ và tên), dob (Ngày sinh), gender (Giới tính),
 * tx1, tx2, tx3, tx4, tx5 (ĐĐGtx 1..5), gk (ĐĐGgk), ck (ĐĐGck), dtb1 (ĐTBmhk1), dtb2 (ĐTBmhk2), dtbcn (ĐTBmcn)
 */
const SAMPLE_STUDENTS_TEMPLATE = [
  { stt: 1, mdd: '2272825039', name: 'Triệu Hoa Tuệ An', dob: '03/03/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 2, mdd: '2204841038', name: 'Dịp Thị Vân Anh', dob: '22/09/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 9.2, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 3, mdd: '2272824988', name: 'Vũ Thiện Anh', dob: '20/11/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 4, mdd: '2276668138', name: 'Triệu Quốc Ánh', dob: '01/02/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: '', ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 5, mdd: '2237266732', name: 'Vi Hoàng Bách', dob: '03/08/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: '', ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 6, mdd: '2220105814', name: 'Đàm Hiếu Bình', dob: '28/02/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 9.2, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 7, mdd: '2272825022', name: 'Nguyễn Bảo Châu', dob: '31/08/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 8, mdd: '2237243084', name: 'Lương Minh Dũng', dob: '27/09/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 9, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 9, mdd: '2202310777', name: 'Bùi Việt Đức', dob: '07/12/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 10, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 10, mdd: '2234630693', name: 'Vi Chu Hải Hà', dob: '24/05/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 9.7, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 11, mdd: '2272825027', name: 'Trần Mỹ Hạnh', dob: '12/02/2011', gender: 'Nữ', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 12, mdd: '2272824997', name: 'Lã Việt Hoàng', dob: '04/06/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 13, mdd: '2237243063', name: 'Nịnh Quốc Huy', dob: '06/08/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 14, mdd: '6702526852', name: 'Lê Chấn Hưng', dob: '29/08/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8.7, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 15, mdd: '2234607234', name: 'Nịnh Quốc Hưng', dob: '11/09/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 16, mdd: '2234668016', name: 'Nguyễn Bảo Khang', dob: '09/09/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 17, mdd: '3320099402', name: 'Phạm Hoàng Bảo Khanh', dob: '13/11/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 10, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 18, mdd: '2220099403', name: 'Nguyễn Gia Khánh', dob: '28/09/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 19, mdd: '2220016703', name: 'Nông Duy Khánh', dob: '18/12/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 20, mdd: '3420099404', name: 'Phạm Ngọc Khánh', dob: '10/09/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8.7, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 21, mdd: '2220099405', name: 'HỒ DUY KHIÊM', dob: '28/06/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 22, mdd: '2203382185', name: 'Vũ Tùng Lâm', dob: '16/03/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 23, mdd: '2272825025', name: 'Bùi Thanh Mai', dob: '19/10/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8.2, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 24, mdd: '2220016711', name: 'Lộc Phương Nga', dob: '06/06/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 25, mdd: '2220016712', name: 'Đinh Hoàng Nghiệp', dob: '08/11/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 26, mdd: '2237301972', name: 'Bàn Thế Nhân', dob: '10/02/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: '', ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 27, mdd: '3420099410', name: 'Phạm Thị Phương Nhi', dob: '03/04/2011', gender: 'Nữ', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 28, mdd: '2272825002', name: 'Phạm Tô Uyển Nhi', dob: '03/10/2011', gender: 'Nữ', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 29, mdd: '2220099413', name: 'Lê Chấn Phong', dob: '24/12/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 30, mdd: '2276668154', name: 'Lương Bằng Quang', dob: '23/12/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 31, mdd: '2272825034', name: 'Khiếu Minh Quân', dob: '17/05/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 32, mdd: '2237243077', name: 'Bùi Danh Thái', dob: '23/10/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 33, mdd: '2220099414', name: 'Đặng Nhật Thành', dob: '16/04/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: '', ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 34, mdd: '2237301974', name: 'Triệu Thị Phương Thảo', dob: '03/06/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 35, mdd: '2272825000', name: 'Vi Hùng Thắng', dob: '25/12/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 36, mdd: '2260412006', name: 'Phạm Mẫn Thi', dob: '21/02/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 37, mdd: '2237243078', name: 'Nguyễn Duy Thịnh', dob: '25/10/2011', gender: 'Nam', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 9, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 38, mdd: '2220016725', name: 'Nịnh Thanh Thùy', dob: '12/07/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 39, mdd: '2235885913', name: 'Khúc Mạnh Trường', dob: '27/02/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8.7, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 40, mdd: '2237301950', name: 'Hoàng Duy Tùng', dob: '31/03/2011', gender: 'Nam', tx1: 9, tx2: 8, tx3: '', tx4: '', tx5: '', gk: 8.2, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 41, mdd: '2272825040', name: 'Tô Khánh Vân', dob: '20/07/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 42, mdd: '2272825038', name: 'Hoàng Khánh Vy', dob: '25/07/2011', gender: 'Nữ', tx1: 9, tx2: 9, tx3: '', tx4: '', tx5: '', gk: 8, ck: '', dtb1: '', dtb2: '', dtbcn: '' },
  { stt: 43, mdd: '2220807810', name: 'Đặng Thị Xoan', dob: '24/08/2011', gender: 'Nữ', tx1: 9, tx2: 8.5, tx3: '', tx4: '', tx5: '', gk: 7.5, ck: '', dtb1: '', dtb2: '', dtbcn: '' }
];

function cloneTemplateStudents() {
  return JSON.parse(JSON.stringify(SAMPLE_STUDENTS_TEMPLATE));
}

/**
 * 1. DỮ LIỆU MẶC ĐỊNH ÁP DỤNG MẪU MỚI CHO TẤT CẢ CÁC LỚP
 */
const DEFAULT_PORTAL_DATA = {
  schoolName: 'Trường THPT Ba Chẽ',
  schoolYear: '2026 - 2027',
  adminPassword: 'admin123',
  classes: [
    {
      id: '10A1',
      name: '10A1',
      grade: '10',
      pin: '1000',
      published: true,
      teacher: 'Vi Anh Quân',
      semester: 'Giữa kỳ I',
      updatedAt: '12/11/2026',
      students: cloneTemplateStudents()
    },
    {
      id: '10A3',
      name: '10A3',
      grade: '10',
      pin: '3000',
      published: true,
      teacher: 'Vi Anh Quân',
      semester: 'Giữa kỳ I',
      updatedAt: '12/11/2026',
      students: cloneTemplateStudents()
    },
    {
      id: '10A4',
      name: '10A4',
      grade: '10',
      pin: '5000',
      published: true,
      teacher: 'Vi Anh Quân',
      semester: 'Giữa kỳ I',
      updatedAt: '12/11/2026',
      students: cloneTemplateStudents()
    },
    {
      id: '11A1',
      name: '11A1',
      grade: '11',
      pin: '1111',
      published: true,
      teacher: 'Vi Anh Quân',
      semester: 'Giữa kỳ I',
      updatedAt: '12/11/2026',
      students: cloneTemplateStudents()
    },
    {
      id: '11A3',
      name: '11A3',
      grade: '11',
      pin: '3333',
      published: true,
      teacher: 'Vi Anh Quân',
      semester: 'Giữa kỳ I',
      updatedAt: '12/11/2026',
      students: cloneTemplateStudents()
    },
    {
      id: '12A1',
      name: '12A1',
      grade: '12',
      pin: '12A1',
      published: true,
      teacher: 'Vi Anh Quân',
      semester: 'Giữa kỳ I',
      updatedAt: '12/11/2026',
      students: cloneTemplateStudents()
    }
  ]
};

// Biến trạng thái toàn cục của ứng dụng
let portalData = null;
let currentUnlockedClassId = null;
let isAdminAuthenticated = false;
let currentAdminEditingClassId = '10A1';

/**
 * 2. KHỞI TẠO DỮ LIỆU & GIAO DIỆN KHI TẢI TRANG
 */
document.addEventListener('DOMContentLoaded', () => {
  loadPortalData();
  initAuthScreen();
  refreshIcons();
});

function loadPortalData() {
  try {
    const rawV2 = localStorage.getItem(STORAGE_KEY);
    if (rawV2) {
      portalData = JSON.parse(rawV2);
    } else {
      // Nếu có dữ liệu từ phiên bản V1, giữ lại danh sách lớp và mã PIN bạn đã cấu hình
      // nhưng áp dụng toàn bộ biểu mẫu học sinh 15 cột mới cho tất cả các lớp
      const rawV1 = localStorage.getItem(LEGACY_STORAGE_KEY);
      if (rawV1) {
        const legacyData = JSON.parse(rawV1);
        portalData = {
          schoolName: legacyData.schoolName || DEFAULT_PORTAL_DATA.schoolName,
          schoolYear: legacyData.schoolYear || DEFAULT_PORTAL_DATA.schoolYear,
          adminPassword: legacyData.adminPassword || DEFAULT_PORTAL_DATA.adminPassword,
          classes: (legacyData.classes || DEFAULT_PORTAL_DATA.classes).map((c) => ({
            id: c.id,
            name: c.name,
            grade: c.grade,
            pin: c.pin,
            published: c.published !== undefined ? c.published : true,
            teacher: 'Vi Anh Quân',
            semester: VALID_SEMESTERS.includes(c.semester) ? c.semester : 'Giữa kỳ I',
            updatedAt: new Date().toLocaleDateString('vi-VN'),
            students: cloneTemplateStudents()
          }))
        };
      } else {
        portalData = JSON.parse(JSON.stringify(DEFAULT_PORTAL_DATA));
      }
      savePortalData();
    }
  } catch (err) {
    console.error('Lỗi đọc dữ liệu LocalStorage, sử dụng dữ liệu mẫu:', err);
    portalData = JSON.parse(JSON.stringify(DEFAULT_PORTAL_DATA));
  }
}

function savePortalData() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(portalData));
  } catch (err) {
    console.error('Không thể lưu vào LocalStorage:', err);
  }
}

function refreshIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

/**
 * 3. HÀM TÍNH ĐIỂM TRUNG BÌNH MÔN HỌC KỲ & ĐIỂM THAM CHIẾU THỐNG KÊ
 * - Công thức Thông tư 22 Bộ GD&ĐT:
 *   ĐTBmhk = (Tổng các cột ĐĐGtx có điểm + 2 * ĐĐGgk + 3 * ĐĐGck) / (Số cột ĐĐGtx có điểm + 5)
 * - Nếu học sinh mới chỉ có điểm ĐĐGtx và ĐĐGgk (đợt Giữa kỳ), hàm getStudentReferenceScore
 *   sẽ tính điểm trung bình tạm thời từ các cột hiện có để phục vụ thống kê nhanh & lọc mức điểm.
 */
function parseScoreOrNull(val) {
  if (val === '' || val === null || val === undefined) return null;
  const normalized = String(val).replace(',', '.').trim();
  if (!normalized) return null;
  const num = parseFloat(normalized);
  return isNaN(num) ? null : num;
}

function computeFormulaSemesterAvg(student) {
  const txCols = [student.tx1, student.tx2, student.tx3, student.tx4, student.tx5]
    .map(parseScoreOrNull)
    .filter((v) => v !== null);
  const gk = parseScoreOrNull(student.gk);
  const ck = parseScoreOrNull(student.ck);

  // Chỉ tính ĐTBmhk chính thức khi đã có ít nhất 1 điểm TX, điểm GK và điểm CK
  if (txCols.length > 0 && gk !== null && ck !== null) {
    const sumTx = txCols.reduce((a, b) => a + b, 0);
    const totalWeight = txCols.length + 5;
    const avg = (sumTx + 2 * gk + 3 * ck) / totalWeight;
    return Math.round(avg * 10) / 10;
  }
  return null;
}

function calculateStudentAverage(student) {
  // Ưu tiên lấy ĐTBmcn -> ĐTBmhk2 -> ĐTBmhk1 nếu đã có trong bảng
  const dtbcn = parseScoreOrNull(student.dtbcn);
  if (dtbcn !== null) return dtbcn;

  const dtb2 = parseScoreOrNull(student.dtb2);
  if (dtb2 !== null) return dtb2;

  const dtb1 = parseScoreOrNull(student.dtb1);
  if (dtb1 !== null) return dtb1;

  // Nếu chưa có cột ĐTBmhk, tính điểm trung bình tạm thời theo trọng số các cột đã nhập (TX hệ số 1, GK hệ số 2, CK hệ số 3)
  const txCols = [student.tx1, student.tx2, student.tx3, student.tx4, student.tx5]
    .map(parseScoreOrNull)
    .filter((v) => v !== null);
  const gk = parseScoreOrNull(student.gk);
  const ck = parseScoreOrNull(student.ck);

  let weightedSum = txCols.reduce((a, b) => a + b, 0);
  let totalWeight = txCols.length;

  if (gk !== null) {
    weightedSum += 2 * gk;
    totalWeight += 2;
  }
  if (ck !== null) {
    weightedSum += 3 * ck;
    totalWeight += 3;
  }

  if (totalWeight === 0) return 0;
  return Math.round((weightedSum / totalWeight) * 10) / 10;
}

/**
 * Chuẩn hóa chuỗi tiếng Việt không dấu để tìm kiếm nhanh
 */
function normalizeVietnamese(str) {
  if (!str) return '';
  return str
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .trim();
}

/**
 * 4. ĐIỀU HƯỚNG GIỮA CÁC MÀN HÌNH (AUTH / SCOREBOARD / ADMIN)
 */
function switchView(viewName) {
  const views = ['view-auth', 'view-scoreboard', 'view-admin-login', 'view-admin-dashboard'];
  views.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.classList.add('hidden');
  });

  const target = document.getElementById(`view-${viewName}`);
  if (target) {
    target.classList.remove('hidden');
  }

  const btnStudent = document.getElementById('nav-btn-student');
  const btnAdmin = document.getElementById('nav-btn-admin');

  if (viewName === 'auth' || viewName === 'scoreboard') {
    btnStudent.className = 'inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-sm font-semibold text-blue-700 bg-blue-50 border border-blue-200/80 transition-all';
    btnAdmin.className = 'inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-sm font-semibold text-slate-700 bg-white hover:bg-slate-100 hover:text-blue-700 border border-slate-200 shadow-sm transition-all';
  } else {
    btnAdmin.className = 'inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-sm font-semibold text-blue-700 bg-blue-50 border border-blue-200/80 transition-all';
    btnStudent.className = 'inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-sm font-semibold text-slate-700 bg-white hover:bg-slate-100 hover:text-blue-700 border border-slate-200 shadow-sm transition-all';
  }

  if (viewName === 'auth') {
    initAuthScreen();
  }

  window.scrollTo({ top: 0, behavior: 'smooth' });
  refreshIcons();
}

/**
 * 5. MÀN HÌNH XÁC THỰC MÃ PIN (TRANG CHỦ)
 */
function initAuthScreen() {
  const gradeSelect = document.getElementById('select-grade');
  const currentGrade = gradeSelect ? gradeSelect.value : '10';
  handleGradeChange(currentGrade);
}

function selectGradeQuick(grade) {
  const gradeSelect = document.getElementById('select-grade');
  if (gradeSelect) {
    gradeSelect.value = grade;
    handleGradeChange(grade);
  }
}

function handleGradeChange(grade) {
  ['10', '11', '12'].forEach((g) => {
    const pill = document.getElementById(`grade-pill-${g}`);
    if (!pill) return;
    if (g === grade) {
      pill.className = 'grade-pill py-2 px-3 rounded-xl text-sm font-bold border border-blue-600 bg-blue-600 text-white shadow-sm shadow-blue-600/20 transition-all flex items-center justify-center gap-1.5';
    } else {
      pill.className = 'grade-pill py-2 px-3 rounded-xl text-sm font-bold border border-slate-200 bg-white text-slate-600 hover:border-blue-300 hover:text-blue-700 transition-all flex items-center justify-center gap-1.5';
    }
  });

  const classSelect = document.getElementById('select-class');
  const classesInGrade = portalData.classes.filter((c) => c.grade === grade);

  classSelect.innerHTML = '';
  classesInGrade.forEach((cls) => {
    const opt = document.createElement('option');
    opt.value = cls.id;
    const statusText = cls.published ? 'Đã công bố điểm' : 'Chưa mở công bố';
    const semText = cls.semester || 'Giữa kỳ I';
    opt.textContent = `Lớp ${cls.name} — (${semText} • ${cls.students.length} học sinh • ${statusText})`;
    classSelect.appendChild(opt);
  });

  handleClassSelectionChange();
  hideAuthError();
}

function handleClassSelectionChange() {
  const classSelect = document.getElementById('select-class');
  const statusBadge = document.getElementById('class-status-badge');
  if (!classSelect || !statusBadge) return;

  const selectedClass = portalData.classes.find((c) => c.id === classSelect.value);
  if (!selectedClass) return;

  if (selectedClass.published) {
    statusBadge.className = 'text-xs font-semibold px-2.5 py-0.5 rounded-md bg-emerald-50 text-emerald-700 border border-emerald-200';
    statusBadge.textContent = '● Đã công bố điểm';
  } else {
    statusBadge.className = 'text-xs font-semibold px-2.5 py-0.5 rounded-md bg-amber-50 text-amber-700 border border-amber-200';
    statusBadge.textContent = '○ Tạm khóa công bố';
  }
  hideAuthError();
}

function togglePasswordVisibility(inputId, iconId) {
  const input = document.getElementById(inputId);
  const icon = document.getElementById(iconId);
  if (!input || !icon) return;

  if (input.type === 'password') {
    input.type = 'text';
    icon.setAttribute('data-lucide', 'eye-off');
  } else {
    input.type = 'password';
    icon.setAttribute('data-lucide', 'eye');
  }
  refreshIcons();
}

function showAuthError(msg) {
  const errBox = document.getElementById('auth-error-box');
  const errMsg = document.getElementById('auth-error-message');
  const authCard = document.getElementById('auth-card');

  errMsg.textContent = msg;
  errBox.classList.remove('hidden');

  authCard.classList.remove('shake-error');
  void authCard.offsetWidth;
  authCard.classList.add('shake-error');
}

function hideAuthError() {
  const errBox = document.getElementById('auth-error-box');
  if (errBox) errBox.classList.add('hidden');
}

function handleUnlockScoreboard(event) {
  event.preventDefault();
  const classId = document.getElementById('select-class').value;
  const enteredPin = document.getElementById('input-class-pin').value.trim();

  const targetClass = portalData.classes.find((c) => c.id === classId);
  if (!targetClass) {
    showAuthError('Không tìm thấy thông tin lớp học đã chọn.');
    return;
  }

  if (!targetClass.published) {
    showAuthError(
      `Bảng điểm lớp ${targetClass.name} hiện đang được giáo viên tạm khóa hoặc chưa đến giờ công bố. Vui lòng quay lại sau!`
    );
    return;
  }

  if (enteredPin !== targetClass.pin) {
    showAuthError(
      `Mã PIN lớp ${targetClass.name} không chính xác! Vui lòng kiểm tra lại mã PIN do giáo viên cung cấp.`
    );
    return;
  }

  hideAuthError();
  currentUnlockedClassId = targetClass.id;
  openScoreboardForClass(targetClass.id);
  showToast(`Mở khóa thành công bảng điểm lớp ${targetClass.name}!`, 'success');
}

/**
 * 6. MÀN HÌNH BẢNG ĐIỂM TOÀN LỚP & THỐNG KÊ NHANH
 */
function openScoreboardForClass(classId) {
  const cls = portalData.classes.find((c) => c.id === classId);
  if (!cls) return;

  currentUnlockedClassId = classId;

  document.getElementById('sb-class-badge').textContent = cls.name;
  document.getElementById('sb-semester-badge').textContent = cls.semester || 'Giữa kỳ I';
  document.getElementById('sb-class-title').textContent = `Bảng Điểm Môn Tin Học — Lớp ${cls.name}`;
  document.getElementById('sb-teacher-name').textContent = cls.teacher || 'Vi Anh Quân';
  document.getElementById('sb-updated-at').textContent = cls.updatedAt || '12/11/2026';

  document.getElementById('print-title').textContent = `BẢNG ĐIỂM HỌC KỲ MÔN TIN HỌC - LỚP ${cls.name}`;
  document.getElementById('print-subtitle').textContent = `Năm học: ${portalData.schoolYear} - Đợt: ${cls.semester} | Môn học: TIN HỌC - GV: ${cls.teacher}`;

  document.getElementById('search-student-input').value = '';
  document.getElementById('filter-tier').value = 'all';
  document.getElementById('sort-by').value = 'stt';

  updateClassStatistics(cls);
  renderScoreboardTable();
  switchView('scoreboard');
}

function updateClassStatistics(cls) {
  const total = cls.students.length;
  document.getElementById('stat-total-students').textContent = total;

  if (total === 0) {
    document.getElementById('stat-class-avg').textContent = '0.0';
    document.getElementById('stat-highest-score').textContent = 'Cao nhất: 0.0';
    document.getElementById('stat-good-count').textContent = '0';
    document.getElementById('stat-good-detail').textContent = '(0 Giỏi, 0 Khá)';
    document.getElementById('stat-pass-rate').textContent = '0%';
    document.getElementById('stat-need-effort').textContent = '0 cần cố gắng';
    return;
  }

  const averages = cls.students.map((s) => calculateStudentAverage(s)).filter((v) => v > 0);
  const validCount = averages.length || 1;
  const sumAvg = averages.reduce((acc, v) => acc + v, 0);
  const classAvg = (sumAvg / validCount).toFixed(1);
  const maxAvg = averages.length > 0 ? Math.max(...averages).toFixed(1) : '0.0';

  const excellentCount = averages.filter((a) => a >= 8.0).length;
  const goodCount = averages.filter((a) => a >= 6.5 && a < 8.0).length;
  const passCount = averages.filter((a) => a >= 5.0).length;
  const effortCount = averages.filter((a) => a < 5.0).length;
  const passRate = Math.round((passCount / validCount) * 100);

  document.getElementById('stat-class-avg').textContent = classAvg;
  document.getElementById('stat-highest-score').textContent = `Cao nhất: ${maxAvg}`;
  document.getElementById('stat-good-count').textContent = excellentCount + goodCount;
  document.getElementById('stat-good-detail').textContent = `(${excellentCount} Giỏi, ${goodCount} Khá)`;
  document.getElementById('stat-pass-rate').textContent = `${passRate}%`;
  document.getElementById('stat-need-effort').textContent = `${effortCount} cần cố gắng`;
}

function getScoreBadgeInfo(avg) {
  if (avg >= 8.0) {
    return {
      label: 'Giỏi / Tốt',
      rowTint: 'hover:bg-emerald-50/40',
      badgeClass: 'bg-emerald-100 text-emerald-800 border-emerald-200',
      scoreTextClass: 'text-emerald-700'
    };
  }
  if (avg >= 6.5) {
    return {
      label: 'Khá',
      rowTint: 'hover:bg-blue-50/40',
      badgeClass: 'bg-blue-100 text-blue-800 border-blue-200',
      scoreTextClass: 'text-blue-700'
    };
  }
  if (avg >= 5.0) {
    return {
      label: 'Đạt',
      rowTint: 'hover:bg-amber-50/40',
      badgeClass: 'bg-amber-100 text-amber-800 border-amber-200',
      scoreTextClass: 'text-amber-700'
    };
  }
  return {
    label: 'Cần cố gắng',
    rowTint: 'hover:bg-rose-50/40',
    badgeClass: 'bg-rose-100 text-rose-800 border-rose-200',
    scoreTextClass: 'text-rose-700'
  };
}

function formatScoreCell(val, isSummaryCol = false) {
  const num = parseScoreOrNull(val);
  if (num === null) return '<span class="text-slate-300">-</span>';

  const display = Number.isInteger(num) ? String(num) : num.toFixed(1);
  if (isSummaryCol) {
    const badge = getScoreBadgeInfo(num);
    return `<span class="inline-flex items-center px-2 py-0.5 rounded-md border font-mono-tech font-extrabold text-xs ${badge.badgeClass}">${display}</span>`;
  }

  if (num >= 8.0) return `<span class="font-mono-tech font-bold text-emerald-700">${display}</span>`;
  if (num < 5.0) return `<span class="font-mono-tech font-bold text-rose-600">${display}</span>`;
  return `<span class="font-mono-tech font-semibold text-slate-700">${display}</span>`;
}

function renderScoreboardTable() {
  const cls = portalData.classes.find((c) => c.id === currentUnlockedClassId);
  if (!cls) return;

  const searchInput = document.getElementById('search-student-input');
  const clearBtn = document.getElementById('btn-clear-search');
  const rawQuery = searchInput ? searchInput.value : '';
  const query = normalizeVietnamese(rawQuery);

  if (clearBtn) {
    clearBtn.classList.toggle('hidden', rawQuery.length === 0);
  }

  const tierFilter = document.getElementById('filter-tier').value;
  const sortBy = document.getElementById('sort-by').value;

  let rows = cls.students.map((student, idx) => {
    const avg = calculateStudentAverage(student);
    return {
      ...student,
      stt: student.stt || idx + 1,
      avg
    };
  });

  if (query) {
    rows = rows.filter(
      (r) =>
        normalizeVietnamese(r.name).includes(query) ||
        normalizeVietnamese(r.mdd || '').includes(query) ||
        normalizeVietnamese(r.dob || '').includes(query)
    );
  }

  if (tierFilter !== 'all') {
    rows = rows.filter((r) => {
      if (tierFilter === 'excellent') return r.avg >= 8.0;
      if (tierFilter === 'good') return r.avg >= 6.5 && r.avg < 8.0;
      if (tierFilter === 'average') return r.avg >= 5.0 && r.avg < 6.5;
      if (tierFilter === 'effort') return r.avg > 0 && r.avg < 5.0;
      return true;
    });
  }

  rows.sort((a, b) => {
    if (sortBy === 'avg-desc') return b.avg - a.avg;
    if (sortBy === 'avg-asc') return a.avg - b.avg;
    if (sortBy === 'name-asc') {
      const lastNameA = a.name.trim().split(/\s+/).pop() || '';
      const lastNameB = b.name.trim().split(/\s+/).pop() || '';
      return lastNameA.localeCompare(lastNameB, 'vi');
    }
    return a.stt - b.stt;
  });

  const tbody = document.getElementById('scoreboard-tbody');
  const emptyState = document.getElementById('scoreboard-empty');
  tbody.innerHTML = '';

  if (rows.length === 0) {
    emptyState.classList.remove('hidden');
    return;
  }
  emptyState.classList.add('hidden');

  rows.forEach((student) => {
    const badge = getScoreBadgeInfo(student.avg);
    const tr = document.createElement('tr');
    tr.className = `transition-colors ${badge.rowTint}`;

    tr.innerHTML = `
      <td class="py-3 px-3 text-center font-mono-tech font-bold text-slate-500 border-r border-slate-100">
        ${student.stt}
      </td>
      <td class="py-3 px-3 text-center font-mono-tech text-xs font-semibold text-slate-600 border-r border-slate-100">
        ${escapeHtml(student.mdd || '')}
      </td>
      <td class="py-3 px-4 font-bold text-slate-900 border-r border-slate-100">
        ${escapeHtml(student.name)}
      </td>
      <td class="py-3 px-3 text-center font-mono-tech text-xs text-slate-600 border-r border-slate-100">
        ${escapeHtml(student.dob || '')}
      </td>
      <td class="py-3 px-2.5 text-center text-xs font-medium text-slate-700 border-r border-slate-100">
        ${escapeHtml(student.gender || '')}
      </td>
      <td class="py-3 px-2 text-center border-r border-slate-100 bg-blue-50/15">${formatScoreCell(student.tx1)}</td>
      <td class="py-3 px-2 text-center border-r border-slate-100 bg-blue-50/15">${formatScoreCell(student.tx2)}</td>
      <td class="py-3 px-2 text-center border-r border-slate-100 bg-blue-50/15">${formatScoreCell(student.tx3)}</td>
      <td class="py-3 px-2 text-center border-r border-slate-100 bg-blue-50/15">${formatScoreCell(student.tx4)}</td>
      <td class="py-3 px-2 text-center border-r border-slate-100 bg-blue-50/15">${formatScoreCell(student.tx5)}</td>
      <td class="py-3 px-3 text-center border-r border-slate-100 font-semibold">${formatScoreCell(student.gk)}</td>
      <td class="py-3 px-3 text-center border-r border-slate-100 font-semibold">${formatScoreCell(student.ck)}</td>
      <td class="py-3 px-3 text-center border-r border-slate-100 bg-blue-50/35">${formatScoreCell(student.dtb1, true)}</td>
      <td class="py-3 px-3 text-center border-r border-slate-100 bg-blue-50/35">${formatScoreCell(student.dtb2, true)}</td>
      <td class="py-3 px-3 text-center bg-indigo-50/40">${formatScoreCell(student.dtbcn, true)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function clearStudentSearch() {
  const input = document.getElementById('search-student-input');
  if (input) {
    input.value = '';
    renderScoreboardTable();
    input.focus();
  }
}

/**
 * 7. TẠO WORKBOOK EXCEL CHUẨN BIỂU MẪU BỘ GD&ĐT / THPT BA CHẼ
 * Dùng chung cho cả "Xuất Excel" và "Tải file Excel mẫu"
 */
function buildStandardSchoolExcelAOA(cls) {
  const teacherName = cls?.teacher || 'Vi Anh Quân';
  const semText = (cls?.semester || 'Giữa kỳ I').includes('II') ? 'II' : 'I';
  const students = cls?.students || SAMPLE_STUDENTS_TEMPLATE;

  const aoa = [
    ['SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ QUẢNG NINH', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['THPT BA CHẼ', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['BẢNG ĐIỂM HỌC KỲ', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    [`Năm học: ${portalData?.schoolYear || '2026-2027'} - Học kỳ: ${semText}`, '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    [`Môn học: TIN HỌC - GV: ${teacherName}`, '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['STT', 'Mã định danh Bộ GD&ĐT', 'Họ và tên', 'Ngày sinh', 'Giới tính', 'ĐĐGtx', '', '', '', '', 'ĐĐGgk', 'ĐĐGck', 'ĐTBmhk1', 'ĐTBmhk2', 'ĐTBmcn'],
    ['', '', '', '', '', 1, 2, 3, 4, 5, '', '', '', '', '']
  ];

  students.forEach((s, idx) => {
    aoa.push([
      s.stt || idx + 1,
      s.mdd || '',
      s.name || '',
      s.dob || '',
      s.gender || '',
      s.tx1 ?? '',
      s.tx2 ?? '',
      s.tx3 ?? '',
      s.tx4 ?? '',
      s.tx5 ?? '',
      s.gk ?? '',
      s.ck ?? '',
      s.dtb1 ?? '',
      s.dtb2 ?? '',
      s.dtbcn ?? ''
    ]);
  });

  aoa.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  aoa.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  aoa.push(['', '', '', '', '', '', '', '', '', '', 'Người lập biểu\n(Ký và ghi rõ họ tên)', '', '', '', '']);

  return aoa;
}

function downloadExcelFromAOA(aoa, fileNameBase, sheetName = 'BangDiem') {
  if (window.XLSX) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    // Thiết lập gộp ô (Merge cells) giống hệt biểu mẫu Excel nhà trường
    ws['!merges'] = [
      { s: { r: 6, c: 0 }, e: { r: 7, c: 0 } }, // STT
      { s: { r: 6, c: 1 }, e: { r: 7, c: 1 } }, // Mã định danh Bộ GD&ĐT
      { s: { r: 6, c: 2 }, e: { r: 7, c: 2 } }, // Họ và tên
      { s: { r: 6, c: 3 }, e: { r: 7, c: 3 } }, // Ngày sinh
      { s: { r: 6, c: 4 }, e: { r: 7, c: 4 } }, // Giới tính
      { s: { r: 6, c: 5 }, e: { r: 6, c: 9 } }, // ĐĐGtx (gộp 5 cột 1..5)
      { s: { r: 6, c: 10 }, e: { r: 7, c: 10 } }, // ĐĐGgk
      { s: { r: 6, c: 11 }, e: { r: 7, c: 11 } }, // ĐĐGck
      { s: { r: 6, c: 12 }, e: { r: 7, c: 12 } }, // ĐTBmhk1
      { s: { r: 6, c: 13 }, e: { r: 7, c: 13 } }, // ĐTBmhk2
      { s: { r: 6, c: 14 }, e: { r: 7, c: 14 } }  // ĐTBmcn
    ];
    ws['!cols'] = [
      { wch: 6 },
      { wch: 16 },
      { wch: 24 },
      { wch: 12 },
      { wch: 9 },
      { wch: 6 },
      { wch: 6 },
      { wch: 6 },
      { wch: 6 },
      { wch: 6 },
      { wch: 8 },
      { wch: 8 },
      { wch: 10 },
      { wch: 10 },
      { wch: 10 }
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${fileNameBase}.xlsx`);
  } else {
    // Fallback CSV chuẩn UTF-8 BOM nếu mất mạng
    const csvContent = aoa
      .map((row) =>
        row
          .map((cell) => {
            const s = String(cell ?? '');
            return s.includes(',') || s.includes('"') || s.includes('\n')
              ? `"${s.replace(/"/g, '""')}"`
              : s;
          })
          .join(',')
      )
      .join('\r\n');
    const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${fileNameBase}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }
}

function exportCurrentClassExcel() {
  const cls = portalData.classes.find((c) => c.id === currentUnlockedClassId);
  if (!cls) return;

  const aoa = buildStandardSchoolExcelAOA(cls);
  downloadExcelFromAOA(aoa, `Bang_Diem_Tin_Hoc_Lop_${cls.name}_THPT_Ba_Che`, `Lop_${cls.name}`);
  showToast(`Đã xuất file Excel chuẩn biểu mẫu cho lớp ${cls.name}!`, 'success');
}

function downloadSampleExcelTemplate() {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId) || portalData.classes[0];
  const aoa = buildStandardSchoolExcelAOA(cls);
  downloadExcelFromAOA(aoa, `Mau_Bang_Diem_Hoc_Ky_Tin_Hoc_THPT_Ba_Che`, 'Bang_Diem_Hoc_Ky');
  showToast('Đã tải xuống file Excel mẫu chuẩn THPT Ba Chẽ!', 'info');
}

/**
 * 8. CỔNG GIÁO VIÊN (ADMIN DASHBOARD)
 */
function openAdminPortal() {
  if (isAdminAuthenticated) {
    initAdminDashboard();
    switchView('admin-dashboard');
  } else {
    document.getElementById('admin-password-input').value = '';
    document.getElementById('admin-login-error').classList.add('hidden');
    switchView('admin-login');
  }
}

function fillDemoAdminPassword() {
  const input = document.getElementById('admin-password-input');
  if (input) {
    input.value = portalData.adminPassword;
    showToast('Đã điền mật khẩu quản trị mẫu. Nhấn "Đăng Nhập Admin" để tiếp tục!', 'info');
  }
}

function handleAdminLogin(event) {
  event.preventDefault();
  const entered = document.getElementById('admin-password-input').value.trim();
  const errEl = document.getElementById('admin-login-error');

  if (entered === portalData.adminPassword) {
    isAdminAuthenticated = true;
    errEl.classList.add('hidden');
    initAdminDashboard();
    switchView('admin-dashboard');
    showToast('Đăng nhập Cổng Giáo Viên thành công!', 'success');
  } else {
    errEl.classList.remove('hidden');
  }
}

function handleAdminLogout() {
  isAdminAuthenticated = false;
  switchView('auth');
  showToast('Đã đăng xuất khỏi Cổng Giáo Viên.', 'info');
}

function initAdminDashboard() {
  if (!portalData.classes.some((c) => c.id === currentAdminEditingClassId)) {
    currentAdminEditingClassId = portalData.classes[0]?.id || '10A1';
  }
  renderAdminClassCards();
  populateAdminClassSelector();
  renderAdminScoreEditor();
  refreshIcons();
}

function renderAdminClassCards() {
  const grid = document.getElementById('admin-class-grid');
  const gradeFilter = document.getElementById('admin-filter-grade').value;
  if (!grid) return;

  grid.innerHTML = '';
  const filteredClasses = portalData.classes.filter(
    (c) => gradeFilter === 'all' || c.grade === gradeFilter
  );

  filteredClasses.forEach((cls) => {
    const totalStudents = cls.students.length;
    const validAvgs = cls.students.map((s) => calculateStudentAverage(s)).filter((v) => v > 0);
    const avgClass =
      validAvgs.length > 0
        ? (validAvgs.reduce((sum, v) => sum + v, 0) / validAvgs.length).toFixed(1)
        : '0.0';

    const card = document.createElement('div');
    const isEditingThis = cls.id === currentAdminEditingClassId;
    card.className = `rounded-2xl p-4 bg-white border transition-all ${
      isEditingThis
        ? 'border-blue-600 ring-2 ring-blue-600/15 shadow-md'
        : 'border-slate-200/90 shadow-sm hover:border-blue-300'
    }`;

    card.innerHTML = `
      <div class="flex items-center justify-between gap-2 pb-3 border-b border-slate-100">
        <div class="flex items-center gap-2.5">
          <span class="px-2.5 py-1 rounded-xl bg-blue-700 text-white font-extrabold text-sm">
            ${cls.name}
          </span>
          <div>
            <p class="text-xs font-bold text-slate-800">Khối ${cls.grade} • ${escapeHtml(cls.semester || 'Giữa kỳ I')} • ${totalStudents} HS</p>
            <p class="text-[11px] text-slate-500">Điểm TB lớp: <strong class="text-blue-700">${avgClass}</strong></p>
          </div>
        </div>

        <label class="inline-flex items-center gap-2 cursor-pointer select-none" title="Bật hoặc tắt công bố điểm lớp ${cls.name}">
          <span class="text-xs font-semibold ${cls.published ? 'text-emerald-700' : 'text-slate-400'}">
            ${cls.published ? 'Đã công bố' : 'Tạm khóa'}
          </span>
          <input
            type="checkbox"
            ${cls.published ? 'checked' : ''}
            onchange="toggleClassPublishStatus('${cls.id}', this.checked)"
            class="sr-only peer"
          />
          <div class="w-10 h-5 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all relative peer-checked:bg-emerald-600"></div>
        </label>
      </div>

      <div class="mt-3">
        <label class="block text-[11px] font-bold uppercase text-slate-500 mb-1">
          Mã PIN mở khóa lớp ${cls.name}
        </label>
        <div class="flex items-center gap-2">
          <input
            type="text"
            id="pin-edit-${cls.id}"
            value="${escapeHtml(cls.pin)}"
            class="w-full rounded-xl border border-slate-300 px-3 py-1.5 text-xs sm:text-sm font-mono-tech font-bold text-slate-900 focus:border-blue-600 focus:outline-none"
          />
          <button
            type="button"
            onclick="saveClassPin('${cls.id}')"
            class="px-3 py-1.5 rounded-xl bg-slate-900 hover:bg-blue-700 text-white text-xs font-bold shrink-0 transition-colors"
          >
            Lưu PIN
          </button>
        </div>
      </div>

      <div class="mt-3.5 pt-2.5 border-t border-slate-100 flex items-center justify-between gap-2">
        <button
          type="button"
          onclick="selectClassForAdminEdit('${cls.id}')"
          class="flex-1 py-1.5 px-3 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1.5 ${
            isEditingThis
              ? 'bg-blue-600 text-white'
              : 'bg-blue-50 text-blue-700 hover:bg-blue-100'
          }"
        >
          <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
          <span>${isEditingThis ? 'Đang chỉnh sửa điểm' : 'Quản lý bảng điểm'}</span>
        </button>

        <button
          type="button"
          onclick="deleteClassInAdmin('${cls.id}')"
          class="p-1.5 rounded-xl text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors"
          title="Xóa lớp ${cls.name}"
        >
          <i data-lucide="trash-2" class="w-4 h-4"></i>
        </button>
      </div>
    `;

    grid.appendChild(card);
  });

  refreshIcons();
}

function toggleClassPublishStatus(classId, isPublished) {
  const cls = portalData.classes.find((c) => c.id === classId);
  if (!cls) return;

  cls.published = isPublished;
  savePortalData();
  renderAdminClassCards();

  showToast(
    isPublished
      ? `Đã BẬT công bố điểm cho lớp ${cls.name}!`
      : `Đã TẠM KHÓA công bố điểm lớp ${cls.name}.`,
    isPublished ? 'success' : 'info'
  );
}

function saveClassPin(classId) {
  const input = document.getElementById(`pin-edit-${classId}`);
  if (!input) return;

  const newPin = input.value.trim();
  if (!newPin) {
    showToast('Mã PIN không được để trống!', 'error');
    return;
  }

  const cls = portalData.classes.find((c) => c.id === classId);
  if (!cls) return;

  cls.pin = newPin;
  savePortalData();
  renderAdminClassCards();
  showToast(`Đã cập nhật Mã PIN lớp ${cls.name} thành: ${newPin}`, 'success');
}

function populateAdminClassSelector() {
  const select = document.getElementById('admin-select-class-editor');
  if (!select) return;

  select.innerHTML = '';
  portalData.classes.forEach((cls) => {
    const opt = document.createElement('option');
    opt.value = cls.id;
    opt.textContent = `${cls.name} (Khối ${cls.grade} • ${cls.students.length} HS)`;
    if (cls.id === currentAdminEditingClassId) opt.selected = true;
    select.appendChild(opt);
  });
}

function selectClassForAdminEdit(classId) {
  currentAdminEditingClassId = classId;
  renderAdminClassCards();
  populateAdminClassSelector();
  renderAdminScoreEditor();

  const editorSection = document.getElementById('admin-score-editor-section');
  if (editorSection) {
    editorSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function handleAdminSelectClassForEdit(classId) {
  currentAdminEditingClassId = classId;
  renderAdminClassCards();
  renderAdminScoreEditor();
}

/**
 * Render bảng chỉnh sửa điểm trực tiếp 15 cột chuẩn biểu mẫu Bộ GD&ĐT
 */
function renderAdminScoreEditor() {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls) return;

  document.getElementById('admin-edit-teacher').value = cls.teacher || 'Vi Anh Quân';
  const semSelect = document.getElementById('admin-edit-semester');
  if (semSelect) {
    semSelect.value = VALID_SEMESTERS.includes(cls.semester) ? cls.semester : 'Giữa kỳ I';
  }

  const tbody = document.getElementById('admin-editor-tbody');
  tbody.innerHTML = '';

  const scoreFields = ['tx1', 'tx2', 'tx3', 'tx4', 'tx5', 'gk', 'ck', 'dtb1', 'dtb2', 'dtbcn'];

  cls.students.forEach((student, idx) => {
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-slate-50/80 transition-colors';

    const scoreInputsHtml = scoreFields
      .map((field) => {
        const isDtb = field.startsWith('dtb');
        const val = student[field] !== undefined && student[field] !== null ? student[field] : '';
        return `
          <td class="py-2 px-1 text-center ${isDtb ? 'bg-blue-50/35' : ''}">
            <input
              type="text"
              inputmode="decimal"
              value="${escapeHtml(val)}"
              placeholder="-"
              onchange="updateStudentField(${idx}, '${field}', this.value)"
              class="w-12 text-center py-1 px-1 rounded-lg border ${
                isDtb ? 'border-blue-300 font-extrabold text-blue-900 bg-white' : 'border-slate-200 font-semibold text-slate-800'
              } font-mono-tech text-xs focus:border-blue-600 focus:outline-none"
            />
          </td>
        `;
      })
      .join('');

    tr.innerHTML = `
      <td class="py-2 px-2 text-center font-mono-tech font-bold text-slate-500 text-xs">
        ${idx + 1}
      </td>
      <td class="py-2 px-1.5">
        <input
          type="text"
          value="${escapeHtml(student.mdd || '')}"
          onchange="updateStudentField(${idx}, 'mdd', this.value)"
          class="w-full rounded-lg border border-slate-200 px-2 py-1 text-xs font-mono-tech text-center text-slate-700 focus:border-blue-600 focus:outline-none"
        />
      </td>
      <td class="py-2 px-2">
        <input
          type="text"
          value="${escapeHtml(student.name || '')}"
          onchange="updateStudentField(${idx}, 'name', this.value)"
          class="w-full rounded-lg border border-slate-200 px-2.5 py-1 text-xs sm:text-sm font-semibold text-slate-900 focus:border-blue-600 focus:outline-none"
        />
      </td>
      <td class="py-2 px-1.5">
        <input
          type="text"
          value="${escapeHtml(student.dob || '')}"
          placeholder="DD/MM/YYYY"
          onchange="updateStudentField(${idx}, 'dob', this.value)"
          class="w-full rounded-lg border border-slate-200 px-2 py-1 text-xs font-mono-tech text-center text-slate-700 focus:border-blue-600 focus:outline-none"
        />
      </td>
      <td class="py-2 px-1.5">
        <select
          onchange="updateStudentField(${idx}, 'gender', this.value)"
          class="w-full rounded-lg border border-slate-200 px-1.5 py-1 text-xs font-medium text-slate-700 focus:border-blue-600 focus:outline-none"
        >
          <option value="Nam" ${student.gender === 'Nam' ? 'selected' : ''}>Nam</option>
          <option value="Nữ" ${student.gender === 'Nữ' ? 'selected' : ''}>Nữ</option>
        </select>
      </td>
      ${scoreInputsHtml}
      <td class="py-2 px-2 text-center">
        <button
          type="button"
          onclick="removeStudentRowInAdmin(${idx})"
          class="p-1.5 rounded-lg text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors"
          title="Xóa học sinh này"
        >
          <i data-lucide="trash" class="w-4 h-4"></i>
        </button>
      </td>
    `;

    tbody.appendChild(tr);
  });

  refreshIcons();
}

function updateStudentField(studentIndex, field, rawValue) {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls || !cls.students[studentIndex]) return;

  const scoreFields = ['tx1', 'tx2', 'tx3', 'tx4', 'tx5', 'gk', 'ck', 'dtb1', 'dtb2', 'dtbcn'];
  if (scoreFields.includes(field)) {
    const trimmed = String(rawValue).trim();
    if (trimmed === '') {
      cls.students[studentIndex][field] = '';
    } else {
      let num = parseFloat(trimmed.replace(',', '.'));
      if (isNaN(num)) {
        cls.students[studentIndex][field] = '';
      } else {
        num = Math.max(0, Math.min(10, Math.round(num * 10) / 10));
        cls.students[studentIndex][field] = num;
      }
    }
  } else {
    cls.students[studentIndex][field] = String(rawValue).trim();
  }

  cls.updatedAt = new Date().toLocaleDateString('vi-VN');
  savePortalData();

  renderAdminScoreEditor();
  renderAdminClassCards();
  showToast(`Đã lưu thay đổi điểm lớp ${cls.name}`, 'success');
}

function saveClassMetadataFromAdmin() {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls) return;

  cls.teacher = document.getElementById('admin-edit-teacher').value.trim();
  cls.semester = document.getElementById('admin-edit-semester').value.trim();
  cls.updatedAt = new Date().toLocaleDateString('vi-VN');

  savePortalData();
  renderAdminClassCards();
  showToast(`Đã cập nhật thông tin lớp ${cls.name} (${cls.semester})`, 'success');
}

function applySemesterToAllClasses() {
  const semEl = document.getElementById('admin-edit-semester');
  if (!semEl) return;
  const selectedSemester = semEl.value;
  portalData.classes.forEach((c) => {
    c.semester = selectedSemester;
  });
  savePortalData();
  renderAdminClassCards();
  showToast(`Đã áp dụng đợt "${selectedSemester}" cho toàn bộ các lớp!`, 'success');
}

function addStudentRowInAdmin() {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls) return;

  const newStt = cls.students.length + 1;
  cls.students.push({
    stt: newStt,
    mdd: '',
    name: `Học sinh mới ${newStt}`,
    dob: '01/01/2011',
    gender: 'Nam',
    tx1: 9,
    tx2: 9,
    tx3: '',
    tx4: '',
    tx5: '',
    gk: 8.5,
    ck: '',
    dtb1: '',
    dtb2: '',
    dtbcn: ''
  });

  cls.updatedAt = new Date().toLocaleDateString('vi-VN');
  savePortalData();
  renderAdminScoreEditor();
  renderAdminClassCards();
  populateAdminClassSelector();
  showToast(`Đã thêm dòng học sinh mới vào lớp ${cls.name}`, 'success');
}

function removeStudentRowInAdmin(studentIndex) {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls) return;

  const targetName = cls.students[studentIndex]?.name || 'học sinh này';
  if (!confirm(`Bạn có chắc muốn xóa "${targetName}" khỏi bảng điểm lớp ${cls.name}?`)) {
    return;
  }

  cls.students.splice(studentIndex, 1);
  cls.students.forEach((s, i) => {
    s.stt = i + 1;
  });

  savePortalData();
  renderAdminScoreEditor();
  renderAdminClassCards();
  populateAdminClassSelector();
  showToast(`Đã xóa học sinh khỏi lớp ${cls.name}`, 'info');
}

/**
 * Tự động tính ĐTBmhk1 (hoặc ĐTBmhk2) từ ĐĐGtx, ĐĐGgk, ĐĐGck
 */
function autoCalculateSemesterAverageForClass() {
  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls) return;

  const isSem2 = (cls.semester || '').includes('II');
  let countUpdated = 0;

  cls.students.forEach((s) => {
    const computed = computeFormulaSemesterAvg(s);
    if (computed !== null) {
      if (isSem2) {
        s.dtb2 = computed;
      } else {
        s.dtb1 = computed;
      }
      countUpdated++;
    } else {
      // Nếu chưa có điểm Cuối kỳ (CK), tính tạm vào ĐTBmhk nếu giáo viên muốn hoặc giữ nguyên
      const refAvg = calculateStudentAverage(s);
      if (refAvg > 0) {
        if (isSem2) s.dtb2 = refAvg;
        else s.dtb1 = refAvg;
        countUpdated++;
      }
    }

    const d1 = parseScoreOrNull(s.dtb1);
    const d2 = parseScoreOrNull(s.dtb2);
    if (d1 !== null && d2 !== null) {
      s.dtbcn = Math.round(((d1 + 2 * d2) / 3) * 10) / 10;
    }
  });

  savePortalData();
  renderAdminScoreEditor();
  renderAdminClassCards();
  showToast(`Đã tự động tính Điểm TB học kỳ cho ${countUpdated} học sinh lớp ${cls.name}!`, 'success');
}

function previewClassAsStudent() {
  openScoreboardForClass(currentAdminEditingClassId);
}

/**
 * 9. IMPORT DỮ LIỆU ĐIỂM TỪ FILE EXCEL (.XLSX / .XLS / .CSV) THEO ĐÚNG MẪU NHÀ TRƯỜNG
 * Nhận diện tự động biểu mẫu có 8 dòng tiêu đề đầu của Sở GD&ĐT Quảng Ninh - THPT Ba Chẽ
 */
function cleanCellScore(val) {
  if (val === '' || val === null || val === undefined) return '';
  const num = parseFloat(String(val).replace(',', '.').trim());
  if (isNaN(num)) return '';
  return Math.max(0, Math.min(10, Math.round(num * 10) / 10));
}

function handleExcelFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const cls = portalData.classes.find((c) => c.id === currentAdminEditingClassId);
  if (!cls) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Đọc toàn bộ sheet dưới dạng mảng 2 chiều (AOA) để giữ nguyên cấu trúc cột gộp của mẫu nhà trường
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

      if (!rows || rows.length === 0) {
        showToast('File Excel không có dữ liệu!', 'error');
        return;
      }

      // Tự động trích xuất tên Giáo viên ở dòng 5 nếu có (VD: "Môn học: TIN HỌC - GV: Vi Anh Quân")
      for (let i = 0; i < Math.min(8, rows.length); i++) {
        const firstCell = String(rows[i][0] || '').trim();
        if (firstCell.includes('GV:')) {
          const parts = firstCell.split('GV:');
          if (parts[1] && parts[1].trim()) {
            cls.teacher = parts[1].trim();
          }
        }
      }

      // Tìm dòng bắt đầu chứa dữ liệu học sinh (Dòng có cột 0 là số STT >= 1 và cột 2 là Họ và tên)
      const importedStudents = [];
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!row || row.length < 3) continue;

        const sttCandidate = String(row[0] ?? '').trim();
        const mddCandidate = String(row[1] ?? '').trim();
        const nameCandidate = String(row[2] ?? '').trim();

        // Bỏ qua dòng tiêu đề ("STT", "Người lập biểu", hoặc dòng phụ 1,2,3,4,5)
        if (!nameCandidate || normalizeVietnamese(nameCandidate) === 'ho va ten') continue;
        if (sttCandidate !== '' && isNaN(parseInt(sttCandidate, 10))) continue;
        // Nếu cả STT và Mã định danh đều trống thì bỏ qua (tránh đọc dòng chữ ký cuối bảng)
        if (sttCandidate === '' && mddCandidate === '') continue;

        importedStudents.push({
          stt: importedStudents.length + 1,
          mdd: mddCandidate,
          name: nameCandidate,
          dob: String(row[3] ?? '').trim(),
          gender: String(row[4] ?? '').trim() || 'Nam',
          tx1: cleanCellScore(row[5]),
          tx2: cleanCellScore(row[6]),
          tx3: cleanCellScore(row[7]),
          tx4: cleanCellScore(row[8]),
          tx5: cleanCellScore(row[9]),
          gk: cleanCellScore(row[10]),
          ck: cleanCellScore(row[11]),
          dtb1: cleanCellScore(row[12]),
          dtb2: cleanCellScore(row[13]),
          dtbcn: cleanCellScore(row[14])
        });
      }

      if (importedStudents.length === 0) {
        showToast('Không tìm thấy danh sách học sinh hợp lệ trong file Excel!', 'error');
        return;
      }

      cls.students = importedStudents;
      cls.updatedAt = new Date().toLocaleDateString('vi-VN');
      savePortalData();

      renderAdminScoreEditor();
      renderAdminClassCards();
      populateAdminClassSelector();
      showToast(
        `Đã nhập thành công ${importedStudents.length} học sinh theo chuẩn biểu mẫu vào lớp ${cls.name}!`,
        'success'
      );
    } catch (err) {
      console.error('Lỗi đọc file Excel:', err);
      showToast('Không thể đọc file Excel. Vui lòng kiểm tra định dạng .xlsx hoặc .csv!', 'error');
    } finally {
      event.target.value = '';
    }
  };

  reader.readAsArrayBuffer(file);
}

/**
 * 10. THÊM / XÓA LỚP HỌC & KHÔI PHỤC DỮ LIỆU MẪU
 */
function openAddClassModal() {
  document.getElementById('new-class-name').value = '';
  document.getElementById('new-class-pin').value = '';
  document.getElementById('modal-add-class').classList.remove('hidden');
}

function closeAddClassModal() {
  document.getElementById('modal-add-class').classList.add('hidden');
}

function handleCreateNewClass(event) {
  event.preventDefault();
  const grade = document.getElementById('new-class-grade').value;
  const rawName = document.getElementById('new-class-name').value.trim().toUpperCase();
  const pin = document.getElementById('new-class-pin').value.trim();
  const semesterEl = document.getElementById('new-class-semester');
  const semester = semesterEl ? semesterEl.value : 'Giữa kỳ I';
  const teacher = document.getElementById('new-class-teacher').value.trim() || 'Vi Anh Quân';

  if (!rawName || !pin) return;

  if (portalData.classes.some((c) => c.id.toUpperCase() === rawName)) {
    showToast(`Lớp ${rawName} đã tồn tại trong hệ thống!`, 'error');
    return;
  }

  const newClassObj = {
    id: rawName,
    name: rawName,
    grade,
    pin,
    published: true,
    teacher,
    semester,
    updatedAt: new Date().toLocaleDateString('vi-VN'),
    students: cloneTemplateStudents()
  };

  portalData.classes.push(newClassObj);
  currentAdminEditingClassId = rawName;
  savePortalData();

  closeAddClassModal();
  initAdminDashboard();
  showToast(`Đã tạo lớp ${rawName} (Mã PIN: ${pin}) áp dụng mẫu bảng điểm chuẩn!`, 'success');
}

function deleteClassInAdmin(classId) {
  if (portalData.classes.length <= 1) {
    showToast('Hệ thống cần giữ lại ít nhất 1 lớp học!', 'error');
    return;
  }

  if (!confirm(`Bạn có chắc chắn muốn xóa lớp ${classId} và toàn bộ bảng điểm của lớp này?`)) {
    return;
  }

  portalData.classes = portalData.classes.filter((c) => c.id !== classId);
  if (currentAdminEditingClassId === classId) {
    currentAdminEditingClassId = portalData.classes[0].id;
  }

  savePortalData();
  initAdminDashboard();
  showToast(`Đã xóa lớp ${classId}.`, 'info');
}

function resetToDefaultData() {
  if (
    !confirm(
      'Khôi phục toàn bộ dữ liệu mẫu chuẩn THPT Ba Chẽ (43 học sinh/lớp) cho tất cả các lớp?'
    )
  ) {
    return;
  }

  portalData = JSON.parse(JSON.stringify(DEFAULT_PORTAL_DATA));
  currentAdminEditingClassId = '10A1';
  savePortalData();
  initAdminDashboard();
  showToast('Đã khôi phục dữ liệu mẫu chuẩn cho tất cả các lớp!', 'success');
}

/**
 * 11. TIỆN ÍCH HIỂN THỊ THÔNG BÁO (TOAST) & BẢO MẬT CHUỖI HTML
 */
function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  if (!container) return;

  const toast = document.createElement('div');
  const styles = {
    success: 'bg-emerald-900 text-white border-emerald-700',
    error: 'bg-rose-900 text-white border-rose-700',
    info: 'bg-slate-900 text-white border-blue-700'
  };

  toast.className = `pointer-events-auto max-w-sm px-4 py-3 rounded-2xl border shadow-xl text-xs sm:text-sm font-semibold flex items-center gap-2.5 view-section ${
    styles[type] || styles.info
  }`;
  toast.innerHTML = `
    <span class="w-2 h-2 rounded-full ${
      type === 'success' ? 'bg-emerald-400' : type === 'error' ? 'bg-rose-400' : 'bg-sky-400'
    } shrink-0"></span>
    <span>${escapeHtml(message)}</span>
  `;

  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.25s ease';
    setTimeout(() => toast.remove(), 260);
  }, 3400);
}

function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
