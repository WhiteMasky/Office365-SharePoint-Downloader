from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException, StaleElementReferenceException
import time
import os
from PIL import Image
import logging
import sys

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PowerPointCapture:
    def __init__(self):
        self.chrome_options = Options()
        # Set browser window size to 1920x1080
        self.chrome_options.add_argument('--window-size=1920,1080')
        self.chrome_options.add_argument('--start-maximized')
        
        # Other necessary settings
        self.chrome_options.add_argument('--disable-infobars')
        self.chrome_options.add_argument('--disable-notifications')
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.chrome_options.add_experimental_option('useAutomationExtension', False)

    def find_and_click_present_button(self, driver, max_attempts=3):
        """Find and click the Present button in the specified frame"""
        for attempt in range(max_attempts):
            try:
                # Wait for button to appear
                present_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '[aria-label="Present"]'))
                )
                
                # Remove overlaying elements
                driver.execute_script("""
                    const overlays = document.querySelectorAll('[role="dialog"], [class*="overlay"], [class*="modal"]');
                    overlays.forEach(overlay => overlay.remove());
                """)
                
                # Ensure button is visible
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", present_button)
                time.sleep(1)
                
                # Try to click
                try:
                    present_button.click()
                except:
                    driver.execute_script("arguments[0].click();", present_button)
                
                time.sleep(2)
                return True
            except StaleElementReferenceException:
                if attempt < max_attempts - 1:
                    time.sleep(1)
                    continue
            except Exception as e:
                logging.warning(f"Click attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(1)
                    continue
        return False

    def try_click_present_button(self, driver):
        """Try to click the Present button in all possible frames"""
        try:
            # First try in main page
            logging.info("Looking for Present button in main page...")
            if self.find_and_click_present_button(driver):
                return True
            
            # Find all iframes
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if not iframes:
                logging.info("No iframes found")
                return False
            
            # Iterate through each iframe
            for i, iframe in enumerate(iframes):
                try:
                    logging.info(f"Switching to iframe {i + 1}/{len(iframes)}...")
                    driver.switch_to.frame(iframe)
                    
                    if self.find_and_click_present_button(driver):
                        return True
                    
                except Exception as e:
                    logging.warning(f"Error processing iframe {i + 1}: {str(e)}")
                finally:
                    driver.switch_to.default_content()
            
            logging.error("Could not find clickable Present button in any location")
            return False
            
        except Exception as e:
            logging.error(f"Failed to click Present button: {str(e)}")
            return False

    def check_presentation_mode(self, driver):
        """Check if in presentation mode"""
        try:
            time.sleep(2)
            # Check URL change
            if "view=present" in driver.current_url.lower():
                return True
            # Check if fullscreen
            is_fullscreen = driver.execute_script("""
                return document.fullscreenElement !== null || 
                       document.webkitFullscreenElement !== null ||
                       document.mozFullScreenElement !== null ||
                       document.msFullscreenElement !== null;
            """)
            if is_fullscreen:
                return True
            # Check specific elements
            try:
                return driver.find_element(By.CSS_SELECTOR, '[class*="presentationMode"]').is_displayed()
            except:
                return False
        except:
            return False
        
    def capture_slides(self, url, output_folder='slides'):
        driver = None
        try:
            # Create output folder
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            logging.info("Starting browser...")
            driver = webdriver.Chrome(options=self.chrome_options)
            driver.maximize_window()
            
            logging.info("Accessing PowerPoint page...")
            driver.get(url)
            time.sleep(5)  # Wait for page to load
            
            # Try to click Present button
            if not self.try_click_present_button(driver):
                # Save page source for debugging
                with open('page_source.html', 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
                logging.info("Page source saved to page_source.html")
                raise Exception("Could not enter presentation mode")
            
            # Wait for presentation mode to load
            time.sleep(3)
            
            # Start screenshot loop
            slide_count = 0
            last_screenshot = None
            screenshots = []
            consecutive_same_count = 0  # Count of consecutive identical screenshots
            
            logging.info("Starting slide capture...")
            while True:
                # Wait for page to load
                try:
                    WebDriverWait(driver, 5).until(
                        lambda d: d.execute_script('return document.readyState') == 'complete'
                    )
                except:
                    logging.warning("Page load timeout")

                # Capture current page
                screenshot_path = os.path.join(output_folder, f'slide_{slide_count:03d}.png')
                
                # Wait for animations to complete
                time.sleep(2)  # Base wait time
                
                # Check if page is still changing
                old_source = driver.page_source
                time.sleep(0.5)
                if old_source != driver.page_source:
                    time.sleep(1.5)  # Wait longer if page is changing
                
                driver.save_screenshot(screenshot_path)
                
                # Check if identical to previous screenshot
                if last_screenshot:
                    current_img = Image.open(screenshot_path)
                    last_img = Image.open(last_screenshot)
                    
                    if self.images_equal(current_img, last_img):
                        consecutive_same_count += 1
                        logging.info(f"Detected identical screenshot ({consecutive_same_count}/10), trying next slide...")
                        
                        if consecutive_same_count >= 10:  # Need 10 consecutive identical screenshots to confirm last slide
                            logging.info("Confirmed last slide reached")
                            os.remove(screenshot_path)
                            break
                        
                        os.remove(screenshot_path)
                        
                        # Try to move to next slide
                        try:
                            actions = ActionChains(driver)
                            actions.send_keys(Keys.ARROW_RIGHT)
                            actions.pause(0.5)
                            actions.perform()
                            time.sleep(2)  # Wait for animation
                        except Exception as e:
                            logging.warning(f"Failed to send right arrow key: {str(e)}")
                            try:
                                driver.execute_script("""
                                    var event = new KeyboardEvent('keydown', {
                                        'key': 'ArrowRight',
                                        'code': 'ArrowRight',
                                        'keyCode': 39,
                                        'which': 39,
                                        'bubbles': true
                                    });
                                    document.dispatchEvent(event);
                                """)
                                time.sleep(2)
                            except Exception as e:
                                logging.error(f"Failed to simulate keypress: {str(e)}")
                        continue
                    else:
                        consecutive_same_count = 0
                        logging.info(f"Captured slide {slide_count + 1}")
                        last_screenshot = screenshot_path
                        screenshots.append(screenshot_path)
                        slide_count += 1
                else:
                    logging.info(f"Captured slide {slide_count + 1}")
                    last_screenshot = screenshot_path
                    screenshots.append(screenshot_path)
                    slide_count += 1
                
                # Simulate right arrow key
                try:
                    actions = ActionChains(driver)
                    actions.send_keys(Keys.ARROW_RIGHT)
                    actions.pause(0.5)
                    actions.perform()
                    time.sleep(2)  # Wait for animation
                except Exception as e:
                    logging.warning(f"Failed to send right arrow key: {str(e)}")
                    try:
                        driver.execute_script("""
                            var event = new KeyboardEvent('keydown', {
                                'key': 'ArrowRight',
                                'code': 'ArrowRight',
                                'keyCode': 39,
                                'which': 39,
                                'bubbles': true
                            });
                            document.dispatchEvent(event);
                        """)
                        time.sleep(2)
                    except Exception as e:
                        logging.error(f"Failed to simulate keypress: {str(e)}")
            
            # Convert to PDF
            if screenshots:
                self.create_pdf(screenshots, os.path.join(output_folder, 'presentation.pdf'))
                logging.info(f"Presentation saved as PDF: {os.path.join(output_folder, 'presentation.pdf')}")
            
            logging.info(f"Total slides captured: {len(screenshots)}")
            
        except Exception as e:
            logging.error(f"An error occurred: {str(e)}")
            if driver:
                # Save page source for debugging
                with open('error_page_source.html', 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
                logging.info("Error page source saved to error_page_source.html")
        finally:
            if driver:
                driver.quit()
    
    def images_equal(self, img1, img2):
        """Compare if two images are identical"""
        if img1.size != img2.size:
            return False
            
        # Convert to grayscale and resize to reduce computation
        size = (img1.size[0] // 2, img1.size[1] // 2)
        img1_gray = img1.convert('L').resize(size)
        img2_gray = img2.convert('L').resize(size)
        
        # Get image data
        data1 = list(img1_gray.getdata())
        data2 = list(img2_gray.getdata())
        
        # Calculate difference
        total_pixels = len(data1)
        diff_pixels = sum(1 for i in range(total_pixels) if abs(data1[i] - data2[i]) > 25)
        
        # Calculate difference percentage
        diff_percentage = (diff_pixels / total_pixels) * 100
        
        return diff_percentage < 2.0  # Allow 2% difference
    
    def create_pdf(self, image_files, output_file):
        """Combine images into PDF"""
        if not image_files:
            return
            
        logging.info("Generating PDF...")
        images = [Image.open(f) for f in image_files]
        first_image = images[0]
        first_image.save(output_file, "PDF", resolution=100.0, save_all=True, append_images=images[1:])

def main():
    try:
        url = input("Enter SharePoint PowerPoint page URL: ").strip()
        if not url:
            logging.error("URL cannot be empty")
            return
            
        output_folder = input("Enter output folder name (default: slides): ").strip() or 'slides'
        
        capture = PowerPointCapture()
        capture.capture_slides(url, output_folder)
        
    except KeyboardInterrupt:
        logging.info("\nProgram interrupted by user")
        sys.exit(0)
    except Exception as e:
        logging.error(f"Program execution error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 