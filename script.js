// Danh sách ánh xạ mã khoa
const KHOA_MAPPING = {
    "401": "Tài chính (TC)",
    "402": "Kế toán - Kiểm toán (KT-KT)",
    "403": "Quản trị kinh doanh (QTKD)",
    "404": "Công nghệ thông tin & kinh tế số (CNTT&KTS)",
    "405": "Kinh doanh quốc tế (KDQT)",
    "406": "Luật",
    "407": "Kinh tế (KT)",
    "408": "Khoa học dữ liệu (KHDL)",
    "751": "Ngôn ngữ Anh (NN)"
};

// Hàm chuyển đổi tiếng Việt có dấu thành không dấu
function removeVietnameseTones(str) {
    return str
        .normalize('NFD') 
        .replace(/[\u0300-\u036f]/g, '') // Loại bỏ các dấu thanh
        .replace(/đ/g, 'd').replace(/Đ/g, 'D'); // Chữ Đ đặc biệt
}

class SinhVien {
    constructor(hoTen, msv) {
        this.hoTen = String(hoTen).trim();
        this.msv = String(msv).trim().toUpperCase();
        this.khoaHoc = this.msv.substring(0, 2); 
        this.tenKhoa = this.lookupKhoa();
        this.email = this.generateEmail();
    }

    lookupKhoa() {
        const maKhoa = this.msv.substring(3, 6);
        return KHOA_MAPPING[maKhoa] || "Khác";
    }

    generateEmail() {
        // 1. Chuyển tên thành không dấu và viết thường
        let unaccentedName = removeVietnameseTones(this.hoTen);
        const cleanName = unaccentedName.toLowerCase().replace(/\s+/g, ' ').trim();
        const parts = cleanName.split(' '); 
        
        if (parts.length === 0 || cleanName === "") return "";
        if (parts.length === 1) return `${parts[0]}.${this.msv.toLowerCase()}@hvnh.edu.vn`;

        // 2. Lấy tên chính
        const tenChinh = parts.pop(); 
        
        // 3. Lấy chữ cái đầu của họ và đệm
        const hoDem = parts.map(tu => tu.charAt(0)).join('');
        
        // VD: "Nguyễn Thành Đạt" -> "nguyen thanh dat" -> ten: "dat", họ đệm: "nt" -> datnt
        return `${tenChinh}${hoDem}.${this.msv.toLowerCase()}@hvnh.edu.vn`;
    }
}

document.getElementById('uploadExcel').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        renderTable(rows);
    };
    reader.readAsArrayBuffer(file);
});

function renderTable(data) {
    const tbody = document.querySelector('#studentTable tbody');
    tbody.innerHTML = ''; 

    if (data.length === 0) return;

    // --- BƯỚC MỚI: Tự động tìm vị trí cột ---
    let nameIdx = -1;
    let msvIdx = -1;
    const headerRow = data[0]; // Giả sử dòng đầu tiên là tiêu đề

    // Quét dòng tiêu đề để tìm xem cột nào là Họ tên, cột nào là MSV
    for (let i = 0; i < headerRow.length; i++) {
        const colName = String(headerRow[i]).toLowerCase();
        if (colName.includes('họ') || colName.includes('tên')) nameIdx = i;
        if (colName.includes('mã') || colName.includes('msv')) msvIdx = i;
    }

    // Nếu không tìm thấy tiêu đề rõ ràng, gán cứng theo file bạn gửi (Cột 1 là MSV, Cột 2 là Họ tên)
    if (nameIdx === -1) nameIdx = 2; 
    if (msvIdx === -1) msvIdx = 1;

    // --- BẮT ĐẦU ĐỌC DỮ LIỆU TỪ DÒNG SỐ 1 (bỏ qua dòng 0 là tiêu đề) ---
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const hoTen = row[nameIdx];
        const msv = row[msvIdx];

        if (hoTen && msv) {
            const sv = new SinhVien(hoTen, msv);
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="name-text">${sv.hoTen}</td>
                <td><span class="msv-tag">${sv.msv}</span></td>
                <td>Khóa ${sv.khoaHoc}</td>
                <td><span class="khoa-tag">${sv.tenKhoa}</span></td>
                <td class="email-text">${sv.email}</td>
            `;
            tbody.appendChild(tr);
        }
    }

    if (tbody.innerHTML === '') {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-state">Không tìm thấy dữ liệu.</td></tr>';
    }
}