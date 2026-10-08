# 🚀 AUTOMATION TESTING E2E - HỆ THỐNG VĂN PHÒNG ĐIỆN TỬ UTC

> **Đồ án / Bài tập thực hành môn Kiểm thử phần mềm (KTPM)**  
> **Sinh viên thực hiện:** Phạm Công Sỹ  
> **Mã số sinh viên (MSSV):** 6451071066  
> **Repository GitHub:** [https://github.com/6451071066-jpg/KTPM_6451071066_PhamCongSy.git](https://github.com/6451071066-jpg/KTPM_6451071066_PhamCongSy.git)  
> **Trường đại học:** Trường Đại học Giao thông Vận tải (UTC)  
> **Hệ thống kiểm thử (AUT):** [https://vanphongdientu.utc.edu.vn/Login](https://vanphongdientu.utc.edu.vn/Login)  

---

## 📌 1. Giới thiệu tổng quan

Dự án này là bộ kiểm thử tự động toàn diện (End-to-End Automation Testing) dành cho màn hình Xác thực & Đăng nhập của cổng thông tin **Văn phòng điện tử UTC**.

### Điểm nổi bật của dự án:
- **Kiến trúc chuẩn công nghiệp:** Xây dựng theo mô hình **Page Object Model (POM)** tách biệt hoàn toàn giữa cấu hình trình duyệt, đối tượng trang và kịch bản kiểm thử.
- **Bao phủ kịch bản sâu rộng:** **16 Test Cases** bao quát từ Bảo mật giao diện (UI Security), Kiểm tra ràng buộc (Form Validation), Chống tấn công dò mật khẩu (Anti-Brute Force với Captcha), Chuyển hướng SSO/OAuth, Phân tích giá trị biên (BVA), Chịu tải buffer (Stress), cho đến Phòng chống Injection (SQL Injection, XSS) và Điều hướng bàn phím (Keyboard UX).
- **Phát hiện lỗi thực tế (Defect Detection):** Bắt được 01 lỗi thực tế của hệ thống tại kịch bản **TC11** (Lỗi không cắt bỏ khoảng trắng - WhiteSpace Trim Bug).
- **Báo cáo chuyên nghiệp đa nền tảng:**
  - Tích hợp báo cáo đồ họa **Allure Report** với đầy đủ thông số môi trường (Environment properties), phân cấp Severity, Steps và tự động chụp ảnh màn hình (Screenshot) khi test thất bại.
  - Tự động xuất bảng báo cáo thực thi ra **Microsoft Excel** gồm 2 Sheet hoàn chỉnh (Dashboard Thống kê & Chi tiết 16 Test Cases).

---

## 🏗️ 2. Kiến trúc dự án (Page Object Model - POM)

```text
6451071066_PCS/
│
├── base/                              # TẦNG NỀN TẢNG (BASE LAYER)
│   ├── __init__.py
│   └── base_page.py                   # Wrapper class đóng gói WebDriver, WebDriverWait & tiện ích dùng chung
│
├── pages/                             # TẦNG ĐỐI TƯỢNG TRANG (PAGE OBJECT LAYER)
│   ├── __init__.py
│   └── login_page.py                  # Chứa toàn bộ Locators và Actions nghiệp vụ của trang Login UTC
│
├── tests/                             # TẦNG KỊCH BẢN KIỂM THỬ (TEST SCRIPT LAYER)
│   ├── __init__.py
│   ├── conftest.py                    # Quản lý fixture khởi tạo Chrome và hook chụp ảnh khi có lỗi
│   └── test_login_e2e.py              # Bộ 16 kịch bản kiểm thử E2E (TC01 -> TC16)
│
├── conftest.py                        # Fixture quản lý WebDriver và xuất thông số môi trường Allure ở root level
├── pytest.ini                         # Cấu hình Pytest (pythonpath, testpaths, tự động lưu allure-results)
├── requirements.txt                   # Danh sách thư viện phụ thuộc của dự án
├── export_to_excel.py                 # Script tự động xuất báo cáo kiểm thử 16 kịch bản ra Excel
├── BaoCao_Test_Login_UTC.xlsx         # Báo cáo thực thi kiểm thử 2 sheets (Dashboard + Chi tiết 16 TC)
├── allure-results/                    # Dữ liệu kết quả test phục vụ tạo Allure Report
└── README.md                          # Tài liệu hướng dẫn dự án
```

---

## 🛠️ 3. Công nghệ và Thư viện sử dụng

| Công nghệ / Thư viện | Phiên bản | Mục đích sử dụng |
|---|---|---|
| **Python** | `3.13+` | Ngôn ngữ lập trình chính |
| **Selenium WebDriver** | `4.20+` | Tự động hóa thao tác trình duyệt Chrome |
| **PyTest** | `9.1+` | Framework quản lý và thực thi kiểm thử |
| **Webdriver Manager** | `4.0+` | Tự động tải ChromeDriver tương thích với trình duyệt |
| **Allure Framework** | `2.16+` | Tạo báo cáo kiểm thử trực quan với biểu đồ Dashboard & Steps |
| **OpenPyXL** | `3.1+` | Xây dựng và định dạng báo cáo kiểm thử ra file Excel 2 sheets |

---

## 📋 4. Danh sách 16 Kịch bản kiểm thử (Test Cases)

| Mã TC | Phân hệ / Loại kiểm thử | Tên kịch bản | Độ nghiêm trọng | Kết quả thực tế |
|---|---|---|---|---|
| **TC01** | UI / Security | Kiểm tra ô mật khẩu ẩn ký tự (`type='password'`) | Critical | **PASSED** |
| **TC02** | UI Interaction | Kiểm tra bật/tắt checkbox *'Giữ tôi luôn đăng nhập'* | Minor | **PASSED** |
| **TC03** | Form Validation | Bấm Đăng nhập khi để trống cả 2 ô | Normal | **PASSED** |
| **TC04** | Form Validation | Nhập Tên đăng nhập nhưng để trống Mật khẩu | Normal | **PASSED** |
| **TC05** | Form Validation | Để trống Tên đăng nhập nhưng có nhập Mật khẩu | Normal | **PASSED** |
| **TC06** | Authentication | Đăng nhập với tài khoản hoặc mật khẩu không chính xác | Critical | **PASSED** |
| **TC07** | Navigation | Kiểm tra điều hướng liên kết *'Bạn quên mật khẩu đăng nhập ?'* | Normal | **PASSED** |
| **TC08** | SSO / OAuth | Kiểm tra điều hướng nút *'Đăng nhập bằng e-mail UTC'* | Normal | **PASSED** |
| **TC09** | Happy Path | Đăng nhập thành công với tài khoản thật | Blocker | **SKIPPED** *(Chờ TK thật)* |
| **TC10** | Security / Rate Limit | Kích hoạt mã bảo mật (Captcha) sau 3 lần sai liên tiếp | Critical | **PASSED** |
| **TC11** | Input Validation | Nhập toàn ký tự khoảng trắng (Space) | Normal | **FAILED** *(Phát hiện Bug)* |
| **TC12** | Boundary (BVA) | Kiểm tra dữ liệu biên dưới độ dài tối thiểu (1 ký tự) | Minor | **PASSED** |
| **TC13** | Stress / Buffer | Kiểm tra độ bền hệ thống với chuỗi cực dài (500 ký tự) | Normal | **PASSED** |
| **TC14** | Security / SQLi | Kiểm tra phòng chống tấn công SQL Injection (`' OR '1'='1`) | Critical | **PASSED** |
| **TC15** | UX / Keyboard | Kiểm tra gửi form đăng nhập bằng phím `ENTER` thay vì click nút | Normal | **PASSED** |
| **TC16** | Security / XSS | Kiểm tra phòng chống tấn công Cross-Site Scripting (`<script>`) | Critical | **PASSED** |

### 🐞 Chi tiết Bug phát hiện được tại `TC11`:
* **Mô tả:** Hệ thống không tự động loại bỏ khoảng trắng (Trim) trước khi gửi form đăng nhập.
* **Kỳ vọng:** Khi người dùng nhập toàn dấu cách `"   "`, hệ thống phải cắt bỏ và báo lỗi: *"Bạn chưa nhập tên đăng nhập"*.
* **Thực tế:** Hệ thống xem khoảng trắng là chuỗi dữ liệu hợp lệ và gửi request lên Server, sau đó trả về: *"Tài khoản hoặc mật khẩu không đúng."*.

---

## ⚡ 5. Hướng dẫn cài đặt và Chạy kiểm thử

### Bước 1: Clone dự án về máy
```bash
git clone https://github.com/6451071066-jpg/KTPM_6451071066_PhamCongSy.git
cd KTPM_6451071066_PhamCongSy
```

### Bước 2: Cài đặt các thư viện cần thiết
```bash
pip install -r requirements.txt
```

### Bước 3: Chạy kiểm thử tự động
- **Chạy toàn bộ 16 test cases:**
  ```bash
  pytest
  ```
  *(Cấu hình trong `pytest.ini` sẽ tự động hiển thị verbose `-v` và ghi dữ liệu ra `allure-results`)*.

- **Chạy riêng 1 test case cụ thể (ví dụ test SQL Injection):**
  ```bash
  pytest -k "test_tc14_sql_injection_defense"
  ```

- **Chạy riêng nhóm test Bảo mật:**
  ```bash
  pytest -k "security or injection or captcha"
  ```

---

## 📊 6. Tạo và Xem Báo cáo kiểm thử

### 1. Báo cáo trực quan với Allure Report
1. **Chạy test và sinh dữ liệu Allure:**
   ```bash
   pytest
   ```
   Dữ liệu kết quả kiểm thử, ảnh chụp màn hình khi có lỗi và thông số môi trường (`environment.properties`) sẽ được tự động lưu vào thư mục `allure-results/`.

2. **Khởi chạy máy chủ xem Allure Dashboard trên trình duyệt:**
   ```bash
   allure serve allure-results
   ```
   *Dashboard sẽ mở tự động trên trình duyệt, hiển thị tỷ lệ đạt, thời gian thực thi, từng bước kiểm thử (`allure.step`) và ảnh chụp màn hình lỗi đính kèm.*

3. **Hoặc xuất báo cáo tĩnh ra thư mục HTML:**
   ```bash
   allure generate allure-results -o allure-report --clean
   allure open allure-report
   ```

### 2. Báo cáo Microsoft Excel
Chạy lệnh sau để cập nhật lại file báo cáo Excel:
```bash
python export_to_excel.py
```
File **`BaoCao_Test_Login_UTC.xlsx`** sẽ được tạo tại thư mục gốc với 2 Sheet chuyên nghiệp:
- **Sheet 1 (Tổng Quan & Thống Kê):** Bảng Dashboard thống kê số lượng, tỷ lệ đạt/lỗi (14 Passed, 1 Failed, 1 Skipped / Tổng 16 TC), thông số hệ thống và thông tin sinh viên MSSV `6451071066`.
- **Sheet 2 (Chi Tiết Test Cases):** Bảng log ghi nhận chi tiết 11 cột chuẩn đặc tả kiểm thử thực tế cho toàn bộ 16 Test Cases.

---

## 📜 7. Quy chuẩn Commit (Conventional Commits)

Lịch sử Git của dự án được commit bài bản theo từng Test Case:
- `test(ui): [TC01] verify password field masks characters`
- `test(ui): [TC02] verify Keep Me Signed In checkbox toggle`
- `test(validation): [TC03] verify validation errors when both fields are empty`
- `test(validation): [TC04] verify validation error when password field is empty`
- `test(validation): [TC05] verify validation error when username field is empty`
- `test(auth): [TC06] verify login is rejected with invalid username or password`
- `test(nav): [TC07] verify navigation to GetPass forgot password page`
- `test(sso): [TC08] verify redirect to Google OAuth when signing in with UTC email`
- `test(auth): [TC09] add successful login scenario with real account (skip)`
- `test(security): [TC10] verify CAPTCHA is triggered after 3 failed login attempts`
- `test(bug): [TC11] detect defect where system does not trim whitespace characters (Failed)`
- `test(bva): [TC12] verify lower boundary value with 1-character input`
- `test(stress): [TC13] verify system robustness with maximum-length 500-character input`
- `test(security): [TC14] verify defense against SQL injection attacks`
- `test(ux): [TC15] verify login form submission using Enter key`
- `test(security): [TC16] verify protection against Cross-Site Scripting (XSS)`
