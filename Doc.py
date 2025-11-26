import time
import os
import urllib.parse
import requests  # Make sure it is installed: pip install requests
 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    ElementClickInterceptedException,
)
 
# Run Edge with this command before starting the script:
# msedge.exe --remote-debugging-port=9222 --user-data-dir="C:\edge-selenum"
 
# Main folder where all student files will be savedi
# وضع رابط الملف الذي نريد التخزين فيه
BASE_FOLDER = r"C:\Users\Ahmad\Downloads\Doc"
os.makedirs(BASE_FOLDER, exist_ok=True)
 
# Connect to the already opened Edge browser
edge_options = webdriver.EdgeOptions()
edge_options.use_chromium = True
edge_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
 
driver = webdriver.Edge(options=edge_options)
wait = WebDriverWait(driver, 20)
 
print("Edge browser connection is open.")
print("Current page:", driver.current_url)
 
 
# ==========================
# 1) Search for a student from the list
# ==========================
 
def search_student_from_list(student_id: str):
    print(f"\nSearching for student: {student_id}")
 
    # First: wait for the loading overlay to disappear before clicking the search field
    try:
        WebDriverWait(driver, 30).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, "div.blockUI.blockOverlay")
            )
        )
    except TimeoutException:
        print("Overlay still visible after 30 seconds.")
 
    # Get the search input field
    search_input = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//input[@name='studentSearch_input' "
            "or @placeholder='Search Student or Number' "
            "or @aria-label='Search Student or Number']"
        ))
    )
 
    # Try to click on the search input field
    try:
        search_input.click()
    except ElementClickInterceptedException:
        print("Another element or overlay blocked clicking on the search field. Trying JavaScript click.")
        driver.execute_script("arguments[0].click();", search_input)
 
    # Clear the field and type the student ID
    search_input.clear()
    search_input.send_keys(student_id)
    print("The number was entered into the search field.")
 
    print("Waiting for the list results to appear...")
 
    # Wait for the listbox containing the search results
    WebDriverWait(driver, 60).until(
        EC.visibility_of_element_located((By.ID, "studentSearch_listbox"))
    )
 
    # Get the first item in the list
    first_item = wait.until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "#studentSearch_listbox li")
        )
    )
    print("Found the first item in the list. Trying to click it...")
 
    # Try to click the first item
    try:
        first_item.click()
    except ElementClickInterceptedException:
        print("A tooltip or overlay blocked clicking on the first item. Trying JavaScript click.")
 
        try:
            wait.until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, "div[role='tooltip'].k-tooltip")
                )
            )
        except TimeoutException:
            pass
 
        driver.execute_script("arguments[0].click();", first_item)
 
    print("Student was selected from the list.")
    time.sleep(1.0)
 
 
# ==========================
# 2) Select All Programs from dropdown programVersionDropDown
# ==========================
 
def click_all_program_versions():
    """
    For each student:
      - Click the Program Version button (id=programVersionDropDown)
      - From the menu, select "All Programs"
    """
    print("Trying to select 'All Programs'...")
 
    time.sleep(1.5)
 
    # Wait for any loading overlay to disappear
    try:
        WebDriverWait(driver, 10).until(
            EC.invisibility_of_element_located(
                (By.CSS_SELECTOR, "div.blockUI.blockOverlay")
            )
        )
    except TimeoutException:
        print("Overlay did not disappear before opening the program menu. Continuing anyway.")
 
    # 1) Click the Program Version dropdown button
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "programVersionDropDown"))
        )
        try:
            btn.click()
        except ElementClickInterceptedException:
            print("Something blocked clicking on the Program Version button. Trying JavaScript click.")
            driver.execute_script("arguments[0].click();", btn)
 
        print("Program Version button clicked (menu opened).")
        time.sleep(0.8)
 
    except TimeoutException:
        print("Could not find the programVersionDropDown button. Continuing without selecting All Programs.")
        return
    except Exception as e:
        print(f"Error while clicking Program Version button: {e}. Continuing.")
        return
 
    # 2) Select "All Programs" from the menu using JavaScript on the parent k-link k-menu-link
    try:
        text_span = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//span[contains(@class,'k-menu-link-text') and normalize-space()='All Programs']"
            ))
        )
 
        driver.execute_script("""
            var el = arguments[0];
            var parent = el.closest('span.k-link.k-menu-link') || el;
            parent.click();
        """, text_span)
 
        print("'All Programs' was selected from the menu.")
        time.sleep(1.5)
 
    except TimeoutException:
        print("Could not find 'All Programs' in the menu. Continuing without changing it.")
    except Exception as e:
        print(f"Error while selecting 'All Programs': {e}. Continuing.")
 
 
# ==========================
# 3) Check if there are documents in the table
# ==========================
 
def has_documents() -> bool:
    # Wait until the grid container is present
    wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div.cmc-grid"))
    )
 
    time.sleep(1)
 
    # Find table rows (each row is one document)
    rows = driver.find_elements(By.CSS_SELECTOR, "tbody.k-table-tbody tr")
 
    if len(rows) > 0:
        print(f"Number of documents in the table: {len(rows)}")
        return True
 
    # If there are no rows, check if the "no data" cell exists
    try:
        no_data_cell = driver.find_element(By.CSS_SELECTOR, "td.no-data")
        if no_data_cell.is_displayed():
            print("No items to display (no documents found).")
            return False
    except NoSuchElementException:
        print("No rows and no clear 'no-data' message. Considering the table empty.")
        return False
 
    return False
 
 
# ==========================
# 4) Download all attachments (paperclip) for the student
# ==========================
 
def guess_extension(file_url: str, resp: requests.Response) -> str:
    # Try to get extension from the URL path
    parsed = urllib.parse.urlparse(file_url)
    file_name_from_url = os.path.basename(parsed.path)
    _, ext = os.path.splitext(file_name_from_url)
 
    if ext:
        return ext
 
    # If no extension in the URL, try to guess from Content-Type header
    ctype = resp.headers.get("Content-Type", "").lower()
 
    if "pdf" in ctype:
        return ".pdf"
    if "jpeg" in ctype or "jpg" in ctype:
        return ".jpg"
    if "png" in ctype:
        return ".png"
    if "tiff" in ctype:
        return ".tiff"
    if "bmp" in ctype:
        return ".bmp"
 
    # Fallback extension
    return ".bin"
 
 
def download_all_attachments(student_id: str):
    print("Trying to collect all attachments (paperclip icons)...")
 
    paperclip_count = len(driver.find_elements(By.CSS_SELECTOR, "span.fa.fa-paperclip"))
 
    if paperclip_count == 0:
        print("No paperclip icons found in the grid.")
        return
 
    print(f"Number of attachments found: {paperclip_count}")
 
    # Create a folder for this student
    student_folder = os.path.join(BASE_FOLDER, student_id)
    os.makedirs(student_folder, exist_ok=True)
 
    for index in range(paperclip_count):
        print(f"\nAttachment number {index + 1}:")
 
        # Re-locate the paperclips each loop in case the DOM changed
        paperclips = driver.find_elements(By.CSS_SELECTOR, "span.fa.fa-paperclip")
        if index >= len(paperclips):
            print("Number of paperclip icons changed during the loop. Stopping.")
            break
 
        paperclip = paperclips[index]
 
        main_window = driver.current_window_handle
        before_windows = set(driver.window_handles)
 
        # Wait for overlay to disappear before clicking
        try:
            WebDriverWait(driver, 20).until(
                EC.invisibility_of_element_located(
                    (By.CSS_SELECTOR, "div.blockUI.blockOverlay")
                )
            )
        except TimeoutException:
            print("Overlay did not disappear before clicking the paperclip. Trying to click anyway.")
 
        # Click the paperclip to open the document in a new tab
        try:
            paperclip.click()
        except ElementClickInterceptedException:
            print("Overlay blocked clicking the paperclip. Trying JavaScript click.")
            driver.execute_script("arguments[0].click();", paperclip)
 
        print("Paperclip clicked. Waiting for the new tab...")
 
        # Wait for a new window/tab to open
        try:
            WebDriverWait(driver, 10).until(
                lambda d: len(d.window_handles) > len(before_windows)
            )
        except TimeoutException:
            print("No new tab appeared after clicking the paperclip.")
            driver.switch_to.window(main_window)
            continue
 
        new_windows = set(driver.window_handles) - before_windows
        if not new_windows:
            print("Could not detect the new tab.")
            driver.switch_to.window(main_window)
            continue
 
        new_window_handle = new_windows.pop()
        driver.switch_to.window(new_window_handle)
 
        time.sleep(3)
 
        file_url = driver.current_url
        print(f"File URL: {file_url}")
 
        # Use requests with the same cookies as the browser
        session = requests.Session()
        for cookie in driver.get_cookies():
            session.cookies.set(cookie["name"], cookie["value"])
 
        resp = session.get(file_url)
        try:
            resp.raise_for_status()
        except Exception as e:
            print(f"Could not download the file from the URL. Error: {e}")
            driver.close()
            driver.switch_to.window(main_window)
            continue
 
        ext = guess_extension(file_url, resp)
 
        file_name = f"file_{index + 1}{ext}"
        file_path = os.path.join(student_folder, file_name)
 
        with open(file_path, "wb") as f:
            f.write(resp.content)
 
        print(f"File saved to: {file_path}")
 
        driver.close()
        driver.switch_to.window(main_window)
 
 
# ==========================
# 5) Student list
# ==========================
 
STUDENT_IDS = [
 "19010500",
    "19010501",
    "19010502",
    "19010503",
    "19010504",
    "19010505",
    "19010507",
    "19010508",
    "19010509",
    "19010510",
]
 
for sid in STUDENT_IDS:
    print("\n==============================")
    print(f"Student: {sid}")
    print("==============================")
 
    search_student_from_list(sid)
    click_all_program_versions()
 
    if not has_documents():
        print("No documents found. Moving to the next student.")
        continue
 
    download_all_attachments(sid)
 
print("\nFinished processing all student IDs in the list.")
# driver.quit()  # Enable this if you want to close the browser when finished