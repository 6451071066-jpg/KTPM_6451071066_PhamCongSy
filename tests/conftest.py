import sys
import os
import pytest
import allure
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Đảm bảo Python luôn tìm thấy thư mục gốc chứa 'base' và 'pages'
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))


@pytest.fixture(scope="function")
def driver(request):
    """Fixture khởi tạo trình duyệt Chrome trước mỗi test và đóng sau khi hoàn tất.
    Tự động chụp ảnh màn hình đính kèm Allure Report nếu test thất bại.
    """
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    # options.add_argument("--headless=new")  # Bỏ comment nếu muốn chạy ngầm trình duyệt
    
    _driver = webdriver.Chrome(service=service, options=options)
    yield _driver
    
    # Chụp ảnh màn hình đính kèm Allure nếu test thất bại
    if hasattr(request.node, "rep_call") and request.node.rep_call.failed:
        try:
            allure.attach(
                _driver.get_screenshot_as_png(),
                name=f"Screenshot_Failure_{request.node.name}",
                attachment_type=allure.attachment_type.PNG
            )
        except Exception:
            pass

    _driver.quit()


@pytest.hookimpl(tryfirst=True, hookwrapper=True)
def pytest_runtest_makereport(item, call):
    """Hook ghi nhận kết quả test để phục vụ việc chụp ảnh màn hình khi có lỗi."""
    outcome = yield
    rep = outcome.get_result()
    setattr(item, "rep_" + rep.when, rep)


def pytest_sessionfinish(session, exitstatus):
    """Tự động ghi thông tin môi trường kiểm thử vào Allure Report."""
    results_dirs = [
        os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "allure-results")),
        os.path.abspath(os.path.join(os.path.dirname(__file__), "allure-results")),
    ]
    env_content = (
        "Project=KTPM_6451071066_PhamCongSy\n"
        "Student_Name=Pham Cong Sy\n"
        "Student_ID=6451071066\n"
        "Git_Repository=https://github.com/6451071066-jpg/KTPM_6451071066_PhamCongSy.git\n"
        "AUT=https://vanphongdientu.utc.edu.vn/Login\n"
        "Browser=Google Chrome\n"
        "Python_Version=3.13\n"
        "Test_Framework=Pytest + Allure\n"
        "Architecture=Page Object Model (POM)\n"
    )
    for r_dir in results_dirs:
        if os.path.exists(r_dir):
            try:
                with open(os.path.join(r_dir, "environment.properties"), "w", encoding="utf-8") as f:
                    f.write(env_content)
            except Exception:
                pass
