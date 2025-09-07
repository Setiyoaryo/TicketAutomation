import os
import csv
import time
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, ElementClickInterceptedException,
    WebDriverException, StaleElementReferenceException, ElementNotInteractableException
)
from dotenv import load_dotenv
import traceback
from datetime import datetime
import sys
import signal
from pathlib import Path

def setup_logging():
    log_filename = f'automation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    log_path = log_dir / log_filename

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_path, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.getLogger('selenium').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)
    return logging.getLogger(__name__)

logger = setup_logging()

try:
    load_dotenv()
except Exception as e:
    logger.error(f"Failed to load .env file: {e}")
    sys.exit(1)

class Config:
    def __init__(self):
        self._validate_environment()

    def _validate_environment(self):
        required_vars = ['USERNAME', 'PASSWORD', 'LOGIN_URL']
        missing_vars = [var for var in required_vars if not os.getenv(var)]
        if missing_vars:
            raise ValueError(f"Missing .env variables: {missing_vars}")

    USERNAME = os.getenv("USERNAME")
    PASSWORD = os.getenv("PASSWORD")
    LOGIN_URL = os.getenv("LOGIN_URL")
    PROXY_SERVER = os.getenv("PROXY_SERVER")
    MASTER_DATA_FILE = os.getenv("MASTER_DATA_FILE", 'master_data_dp.csv')
    DAILY_INPUT_FILE = os.getenv("DAILY_INPUT_FILE", 'input_dp.txt')

    DEFAULT_TIMEOUT = int(os.getenv("DEFAULT_TIMEOUT", "15"))
    SHORT_TIMEOUT = int(os.getenv("SHORT_TIMEOUT", "5"))
    LONG_TIMEOUT = int(os.getenv("LONG_TIMEOUT", "30"))
    MAX_RETRIES = int(os.getenv("MAX_RETRIES", "3"))
    RETRY_DELAY = int(os.getenv("RETRY_DELAY", "2"))
    DROPDOWN_RETRIES = int(os.getenv("DROPDOWN_RETRIES", "3"))

    SELECTORS = {
        'login_button': ["#App > div > div.col-lg-5.d-flex.align-items-center.justify-content-center > div > form > div.my-3 > button", "button[type='submit']", ".btn-primary"],
        'configuring_menu': ["#sidebar > ul > li:nth-child(3) > ul > li:nth-child(1) > a", "a[href*='configuring']"],
        'dp_menu': ["#menu_0_1 > ul > li:nth-child(4) > a", "a[href*='dp']"],
        'city_input': ["#vs1__combobox > div.vs__selected-options > input", "[id*='vs1'] input", ".vs__search"],
        'rk_input': ["#vs2__combobox > div.vs__selected-options > input", "[id*='vs2'] input"],
        'dp_input': ["#vs3__combobox > div.vs__selected-options > input", "[id*='vs3'] input"],
        'filter_button': ["#dp_comp > div > div > div:nth-child(9) > div > a.btn.btn-primary", ".btn-primary", "button[type='submit']"],
        'data_row': ["#lists_dp > tbody > tr", "tbody tr"],
        'result_dp_code_cell': ["#lists_dp > tbody > tr:first-child > td:nth-child(2)", "tbody tr:first-child td:nth-child(2)"],
        'no_data_message': ["//td[contains(text(),'No data available in table')]", "//td[contains(text(),'No data')]", ".dataTables_empty"],
        'create_ticket_icon': ["#lists_dp > tbody > tr:first-child > td:nth-child(9) > a.btn.btn-success.btn-action > i", "tbody tr:first-child .btn-success i", ".btn-success"],
        'final_create_button': ["#dp_comp > div > div > div.v--modal-overlay.scrollable > div > div.v--modal-box.v--modal > div.modal-body.card.border-gradient-mask2 > div > div > div > div.row.justify-content-center > div > a", ".modal-body .btn-primary", ".v--modal .btn"],
        'confirm_create_button': ["body > div.swal2-container.swal2-center.swal2-backdrop-show > div > div.swal2-actions > button.swal2-confirm.swal2-styled", ".swal2-confirm", ".swal2-styled"],
        'loading_overlay': ["div.vld-background", ".loading", ".overlay"],
        'username_input': ["/html/body/div[1]/div/div/div[1]/div[1]/div/form/div[1]/div/input", "input[type='text']", "input[name='username']", "#username"],
        'password_input': ["/html/body/div[1]/div/div/div[1]/div[1]/div/form/div[2]/div/input", "input[type='password']", "input[name='password']", "#password"]
    }

config = Config()

class WebDriverManager:
    def __init__(self):
        self.driver = None
        self.service = None

    def create_driver(self):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            options.add_argument("--disable-gpu")
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-extensions")
            options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option('useAutomationExtension', False)
            options.add_experimental_option("excludeSwitches", ["enable-automation"])

            if config.PROXY_SERVER:
                options.add_argument(f'--proxy-server={config.PROXY_SERVER}')

            prefs = {
                "profile.default_content_setting_values": {"notifications": 2, "media_stream": 2, "geolocation": 2, "popups": 2},
                "profile.default_content_settings": {"popups": 2}
            }
            options.add_experimental_option("prefs", prefs)

            try:
                self.service = Service()
                self.driver = webdriver.Chrome(service=self.service, options=options)
            except Exception:
                self.driver = webdriver.Chrome(options=options)

            self.driver.implicitly_wait(2)
            self.driver.set_page_load_timeout(config.LONG_TIMEOUT)
            self.driver.set_script_timeout(config.DEFAULT_TIMEOUT)

            self.driver.execute_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
            """)
            return self.driver
        except Exception as e:
            logger.error(f"Failed to create WebDriver: {e}")
            self.cleanup()
            raise

    def cleanup(self):
        try:
            if self.driver:
                try: self.driver.quit()
                except:
                    try: self.driver.close()
                    except: pass
                finally: self.driver = None
            if self.service:
                try: self.service.stop()
                except: pass
                finally: self.service = None
        except Exception as e:
            logger.warning(f"Error during cleanup: {e}")

class ElementHelper:
    def __init__(self, driver):
        self.driver = driver

    def wait_for_element(self, selectors, timeout=None, clickable=True):
        timeout = timeout or config.DEFAULT_TIMEOUT
        if not isinstance(selectors, list): selectors = [selectors]
        for selector in selectors:
            try:
                by = By.XPATH if selector.startswith(('//', './/')) else By.CSS_SELECTOR
                if clickable:
                    element = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((by, selector)))
                else:
                    element = WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by, selector)))
                if element.is_displayed() and element.is_enabled(): return element
            except: continue
        return None

    def safe_click(self, element_or_selectors, element_name="element", use_js=False, max_attempts=3):
        element = self.wait_for_element(element_or_selectors, clickable=True) if isinstance(element_or_selectors, (list, str)) else element_or_selectors
        if not element: return False
        for attempt in range(max_attempts):
            try:
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", element)
                time.sleep(0.5)
                if use_js or attempt > 0:
                    self.driver.execute_script("arguments[0].click();", element)
                else:
                    try: element.click()
                    except ElementClickInterceptedException: self.driver.execute_script("arguments[0].click();", element)
                return True
            except StaleElementReferenceException:
                element = self.wait_for_element(element_or_selectors, clickable=True) if isinstance(element_or_selectors, (list, str)) else None
                if not element: return False
            except: time.sleep(1)
        return False

    def handle_vue_select_dropdown(self, input_selectors, value_to_select, max_retries=None):
        max_retries = max_retries or config.DROPDOWN_RETRIES
        for attempt in range(max_retries):
            try:
                input_field = self.wait_for_element(input_selectors, timeout=10)
                if not input_field: continue
                self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", input_field)
                time.sleep(0.5)
                input_field.click()
                time.sleep(0.3)
                try:
                    input_field.clear()
                    time.sleep(0.2)
                    input_field.send_keys(Keys.CONTROL + "a")
                    time.sleep(0.1)
                    input_field.send_keys(Keys.DELETE)
                    time.sleep(0.2)
                except: pass
                for char in value_to_select:
                    input_field.send_keys(char)
                    time.sleep(0.05)
                time.sleep(1.5)
                exact_match_xpath = f"//ul[contains(@class, 'vs__dropdown-menu')]//li[normalize-space()='{value_to_select}']"
                option = self.wait_for_element([exact_match_xpath], timeout=5)
                if option and self.safe_click(option, f"Option '{value_to_select}'"):
                    time.sleep(0.5)
                    return True
                contains_xpath = f"//ul[contains(@class, 'vs__dropdown-menu')]//li[contains(normalize-space(), '{value_to_select}')]"
                option = self.wait_for_element([contains_xpath], timeout=3)
                if option and self.safe_click(option, f"Option containing '{value_to_select}'"):
                    time.sleep(0.5)
                    return True
                try:
                    input_field.send_keys(Keys.ARROW_DOWN)
                    time.sleep(0.3)
                    input_field.send_keys(Keys.ENTER)
                    time.sleep(0.5)
                    return True
                except: pass
                try: input_field.send_keys(Keys.ESCAPE); time.sleep(0.5)
                except: pass
            except Exception:
                try:
                    input_field = self.wait_for_element(input_selectors, timeout=3)
                    if input_field: input_field.send_keys(Keys.ESCAPE)
                except: pass
                time.sleep(1)
        return False

    def wait_for_page_load(self, timeout=None):
        timeout = timeout or config.DEFAULT_TIMEOUT
        try:
            for selector in config.SELECTORS['loading_overlay']:
                try:
                    WebDriverWait(self.driver, timeout).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, selector)))
                    break
                except: continue
            WebDriverWait(self.driver, timeout).until(lambda driver: driver.execute_script("return document.readyState") == "complete")
            time.sleep(1)
            return True
        except: return False

class DataManager:
    @staticmethod
    def validate_file_exists(filename, file_type="file"):
        if not filename: raise ValueError(f"{file_type} filename cannot be empty")
        file_path = Path(filename)
        if not file_path.exists(): raise FileNotFoundError(f"{file_type} not found: {filename}")
        if not file_path.is_file(): raise ValueError(f"Path is not a file: {filename}")
        return file_path

    @staticmethod
    def load_master_data(filename=None):
        filename = filename or config.MASTER_DATA_FILE
        try:
            file_path = DataManager.validate_file_exists(filename, "Master data file")
            master_data = {}
            required_fields = ['Kode_DP', 'City', 'RK']
            with open(file_path, mode='r', newline='', encoding='utf-8-sig') as file:
                sample = file.read(1024); file.seek(0)
                sniffer = csv.Sniffer()
                try: delimiter = sniffer.sniff(sample).delimiter
                except: delimiter = ','
                reader = csv.DictReader(file, delimiter=delimiter)
                if not reader.fieldnames: raise ValueError("CSV file is empty or invalid")
                cleaned_fieldnames = [field.strip() for field in reader.fieldnames]
                missing_fields = [field for field in required_fields if field not in cleaned_fieldnames]
                if missing_fields: raise ValueError(f"CSV must have columns: {required_fields}. Missing: {missing_fields}")
                for row_num, row in enumerate(reader, start=2):
                    try:
                        kode_dp = row['Kode_DP'].strip() if row.get('Kode_DP') else ''
                        city = row['City'].strip() if row.get('City') else ''
                        rk = row['RK'].strip() if row.get('RK') else ''
                        if not all([kode_dp, city, rk]): continue
                        master_data[kode_dp] = {'City': city, 'RK': rk, 'row_number': row_num}
                    except: continue
            if not master_data: raise ValueError("No valid data found in master data file")
            logger.info(f"Master data loaded successfully. Found {len(master_data)} entries.")
            return master_data
        except Exception as e:
            logger.error(f"Failed to load master data: {e}")
            return None

    @staticmethod
    def read_daily_input(filename=None):
        filename = filename or config.DAILY_INPUT_FILE
        try:
            file_path = DataManager.validate_file_exists(filename, "Daily input file")
            seen_codes = set()
            with open(file_path, mode='r', encoding='utf-8') as file:
                kode_dp_list = []
                for line in file:
                    line = line.strip()
                    if not line or line.startswith('#') or line in seen_codes: continue
                    seen_codes.add(line)
                    kode_dp_list.append(line)
            logger.info(f"Daily work list ready. Found {len(kode_dp_list)} tickets.")
            return kode_dp_list
        except Exception as e:
            logger.error(f"Failed to read daily work list: {e}")
            return []

class AutomationBot:
    def __init__(self):
        self.driver_manager = WebDriverManager()
        self.driver = None
        self.helper = None
        self.processed_dps = set()
        self.stats = {'successful': 0, 'failed': 0, 'skipped': 0, 'start_time': None, 'end_time': None}

    def initialize(self):
        try:
            logger.info("Starting bot...")
            self.driver = self.driver_manager.create_driver()
            if not self.driver: raise RuntimeError("Failed to create WebDriver")
            self.helper = ElementHelper(self.driver)
            self.stats['start_time'] = datetime.now()
            logger.info("Bot is ready.")
            return True
        except Exception as e:
            logger.error(f"Bot initialization failed: {e}")
            self.cleanup()
            return False

    def cleanup(self):
        try:
            if self.stats['start_time'] and not self.stats['end_time']: self.stats['end_time'] = datetime.now()
            if self.driver_manager: self.driver_manager.cleanup()
        except Exception as e:
            logger.warning(f"Error during cleanup: {e}")
    
    def inject_network_listener(self):
        script = """
        (function() {
            if (window.xhr_listener_injected) { return; }
            window.xhr_listener_injected = true;
            window.ticket_creation_response = null;
            var XHR = XMLHttpRequest.prototype;
            var send = XHR.send;
            XHR.send = function() {
                this.addEventListener('load', function() {
                    if (this.responseURL && this.responseURL.includes('/project-management/configuring/dp/create-ticket')) {
                        try {
                            window.ticket_creation_response = JSON.parse(this.responseText);
                        } catch (e) {
                            window.ticket_creation_response = { error: 'JSON parse error', data: this.responseText };
                        }
                    }
                });
                return send.apply(this, arguments);
            };
        })();
        """
        try:
            self.driver.execute_script(script)
        except Exception as e:
            logger.error(f"Failed to inject network listener: {e}")

    def login(self):
        if not all([config.LOGIN_URL, config.USERNAME, config.PASSWORD]):
            logger.error("Login details are incomplete"); return False
        for attempt in range(3):
            try:
                logger.info(f"Login attempt {attempt + 1}...")
                self.driver.get(config.LOGIN_URL)
                if not self.helper.wait_for_page_load(timeout=config.LONG_TIMEOUT): continue
                username_field = self.helper.wait_for_element(config.SELECTORS['username_input'])
                if not username_field: continue
                username_field.clear(); time.sleep(0.2); username_field.send_keys(config.USERNAME); time.sleep(0.3)
                password_field = self.helper.wait_for_element(config.SELECTORS['password_input'])
                if not password_field: continue
                password_field.clear(); time.sleep(0.2); password_field.send_keys(config.PASSWORD); time.sleep(0.3)
                login_button = self.helper.wait_for_element(config.SELECTORS['login_button'])
                if not login_button or not self.helper.safe_click(login_button, "Login button", use_js=True): continue
                try:
                    WebDriverWait(self.driver, config.LONG_TIMEOUT).until(EC.presence_of_element_located((By.ID, "sidebar")))
                    if 'login' not in self.driver.current_url.lower():
                        logger.info("Login successful!")
                        return True
                except: continue
            except Exception:
                if attempt < 2: time.sleep(config.RETRY_DELAY)
        logger.error("Login failed after multiple attempts"); return False

    def navigate_to_dp_menu(self):
        try:
            logger.info("Navigating to ticket creation page...")
            configuring_menu = self.helper.wait_for_element(config.SELECTORS['configuring_menu'])
            if not configuring_menu or not self.helper.safe_click(configuring_menu, "Configuring menu", use_js=True): return False
            time.sleep(2)
            dp_menu = self.helper.wait_for_element(config.SELECTORS['dp_menu'])
            if not dp_menu or not self.helper.safe_click(dp_menu, "DP menu", use_js=True): return False
            if not self.helper.wait_for_page_load(timeout=config.LONG_TIMEOUT): logger.warning("DP page is loading slowly...")
            city_input = self.helper.wait_for_element(config.SELECTORS['city_input'], timeout=5)
            if city_input:
                logger.info("On DP page.")
                self.inject_network_listener()
                return True
            else: return False
        except Exception as e:
            logger.error(f"Navigation failed: {e}"); return False

    def validate_filter_result(self, expected_kode_dp, max_attempts=3):
        for attempt in range(max_attempts):
            try:
                time.sleep(2)
                for no_data_selector in config.SELECTORS['no_data_message']:
                    try:
                        by = By.XPATH if no_data_selector.startswith('//') else By.CSS_SELECTOR
                        if self.driver.find_element(by, no_data_selector).is_displayed(): return False, "NO_DATA"
                    except: continue
                result_cell = self.helper.wait_for_element(config.SELECTORS['result_dp_code_cell'], timeout=10, clickable=False)
                if result_cell:
                    result_dp_text = result_cell.text.strip()
                    if result_dp_text == expected_kode_dp: return True, "MATCH"
                    else: return False, "MISMATCH"
                else:
                    if attempt < max_attempts - 1: time.sleep(1); continue
            except Exception:
                if attempt < max_attempts - 1: time.sleep(1)
        return False, "VALIDATION_FAILED"

    def check_ticket_creation_status(self, timeout=120):
        logger.info("Waiting for ticket creation API response...")
        start_time = time.time()
        end_time = start_time + timeout
        while time.time() < end_time:
            response = self.driver.execute_script("return window.ticket_creation_response;")
            if response:
                duration = time.time() - start_time
                logger.info(f"API response received in {duration:.2f} seconds.")
                if response.get("code") == 200:
                    logger.info(f"Success message: {response.get('message')}")
                    return "SUCCESS"
                else:
                    logger.error(f"API responded with an error code: {response}")
                    return "FAIL"
            time.sleep(0.5)
        logger.error("Timeout: No response from ticket creation API.")
        return "TIMEOUT"

    def process_ticket_creation(self, city, rk, kode_dp_value):
        try:
            if not self.helper.handle_vue_select_dropdown(config.SELECTORS['city_input'], city): return "FAIL"
            time.sleep(0.5)
            if not self.helper.handle_vue_select_dropdown(config.SELECTORS['rk_input'], rk): return "FAIL"
            time.sleep(0.5)
            if not self.helper.handle_vue_select_dropdown(config.SELECTORS['dp_input'], kode_dp_value): return "FAIL"
            time.sleep(1)
            filter_button = self.helper.wait_for_element(config.SELECTORS['filter_button'])
            if not filter_button: return "FAIL"
            validation_passed = False
            for filter_attempt in range(3):
                if not self.helper.safe_click(filter_button, "Filter button", use_js=True): continue
                self.helper.wait_for_page_load()
                validation_result, validation_status = self.validate_filter_result(kode_dp_value)
                if validation_status == "NO_DATA": return "FAIL"
                elif validation_status == "MATCH": validation_passed = True; break
                elif validation_status == "MISMATCH": time.sleep(1); continue
                else: time.sleep(1); continue
            if not validation_passed: return "FAIL"
            create_ticket_icon = self.helper.wait_for_element(config.SELECTORS['create_ticket_icon'])
            if not create_ticket_icon or not self.helper.safe_click(create_ticket_icon, "Create Ticket Icon", use_js=True): return "FAIL"
            self.helper.wait_for_page_load(); time.sleep(2)
            final_create_button = self.helper.wait_for_element(config.SELECTORS['final_create_button'])
            if not final_create_button or not self.helper.safe_click(final_create_button, "Final Create Button", use_js=True): return "FAIL"
            time.sleep(1)
            self.driver.execute_script("window.ticket_creation_response = null;")
            confirm_button = self.helper.wait_for_element(config.SELECTORS['confirm_create_button'])
            if not confirm_button or not self.helper.safe_click(confirm_button, "Confirmation Button", use_js=True): return "FAIL"
            return self.check_ticket_creation_status()
        except Exception as e:
            logger.error(f"Failed to create ticket for '{kode_dp_value}': {e}")
            return "FAIL"

    def handle_page_refresh_and_navigation(self):
        try:
            self.driver.refresh()
            time.sleep(config.RETRY_DELAY)
            self.helper.wait_for_page_load(timeout=config.LONG_TIMEOUT)
            return self.navigate_to_dp_menu()
        except Exception as e:
            logger.error(f"Failed during refresh: {e}"); return False

    def generate_final_report(self, total_items):
        try:
            self.stats['end_time'] = datetime.now()
            duration = self.stats['end_time'] - self.stats['start_time']
            logger.info("="*50)
            logger.info("Automation Complete")
            logger.info("="*50)
            logger.info(f"Total duration: {str(duration).split('.')[0]}")
            logger.info(f"Total: {total_items} | Successful: {self.stats['successful']} | Failed: {self.stats['failed']} | Skipped: {self.stats['skipped']}")
            new_items = total_items - self.stats['skipped']
            if new_items > 0:
                success_rate = (self.stats['successful'] / new_items * 100)
                logger.info(f"Success rate: {success_rate:.1f}%")
            if duration.total_seconds() > 0:
                items_per_minute = (self.stats['successful'] / (duration.total_seconds() / 60))
                logger.info(f"Speed: {items_per_minute:.1f} tickets/minute")
            logger.info("="*50)
        except Exception as e:
            logger.error(f"Failed to generate report: {e}")

    def run_automation(self):
        try:
            logger.info("Starting DP Ticket Automation - By Setiyo Aryo")
            master_data = DataManager.load_master_data()
            if not master_data: return False
            kode_dp_list = DataManager.read_daily_input()
            if not kode_dp_list: return False
            if not self.login(): return False
            if not self.navigate_to_dp_menu(): return False
            self.processed_dps.clear()
            total_items = len(kode_dp_list)
            logger.info(f"Processing {total_items} tickets...")
            for index, kode_dp_value in enumerate(kode_dp_list, 1):
                try:
                    if kode_dp_value in self.processed_dps:
                        logger.info(f"[{index}/{total_items}] Skipping: {kode_dp_value} (already processed)"); self.stats['skipped'] += 1; continue
                    if kode_dp_value not in master_data:
                        logger.warning(f"[{index}/{total_items}] Skipping: {kode_dp_value} (not in master data)"); self.stats['skipped'] += 1; continue
                    data = master_data[kode_dp_value]
                    city, rk = data['City'], data['RK']
                    logger.info(f"[{index}/{total_items}] Processing: {kode_dp_value}")
                    
                    final_status = "FAIL"
                    for attempt in range(config.MAX_RETRIES):
                        try:
                            result = self.process_ticket_creation(city, rk, kode_dp_value)
                            if result == "SUCCESS":
                                final_status = "SUCCESS"
                                break
                            if result == "TIMEOUT":
                                final_status = "TIMEOUT"
                                break
                            
                            if attempt < config.MAX_RETRIES - 1:
                                logger.warning(f"Attempt {attempt + 1} failed, retrying...")
                                if not self.handle_page_refresh_and_navigation(): break
                                time.sleep(config.RETRY_DELAY)
                        except Exception as e:
                            logger.error(f"Error on attempt {attempt + 1}: {str(e)[:50]}")
                            if attempt < config.MAX_RETRIES - 1: time.sleep(config.RETRY_DELAY)
                    
                    if final_status == "SUCCESS":
                        self.processed_dps.add(kode_dp_value); self.stats['successful'] += 1; logger.info(f"SUCCESS: {kode_dp_value}")
                    elif final_status == "TIMEOUT":
                        self.stats['skipped'] += 1; logger.warning(f"SKIPPED: {kode_dp_value} (API response timeout). Coba cek di ticket list intranet untuk kode dp {kode_dp_value} muncul atau tidak")
                    else:
                        self.stats['failed'] += 1; logger.error(f"FAILED: {kode_dp_value}")
                    
                    time.sleep(1)
                except KeyboardInterrupt:
                    logger.info("Process interrupted by user"); break
                except Exception as e:
                    logger.error(f"Error while processing {kode_dp_value}: {str(e)[:50]}"); self.stats['failed'] += 1
            self.generate_final_report(total_items)
            return True
        except Exception as e:
            logger.error(f"Automation failed with a fatal error: {e}"); return False
        finally:
            self.cleanup()

def signal_handler(signum, frame):
    logger.info("Interrupt signal received, cleaning up..."); sys.exit(0)

def main():
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    bot = AutomationBot()
    try:
        if not bot.initialize(): return False
        return bot.run_automation()
    except KeyboardInterrupt:
        logger.info("Process interrupted"); return False
    except Exception as e:
        logger.error(f"Main execution error: {e}"); return False
    finally:
        bot.cleanup()

if __name__ == "__main__":
    try:
        success = main()
        sys.exit(0 if success else 1)
    except Exception as e:
        logger.error(f"Unhandled fatal error: {e}"); sys.exit(1)
