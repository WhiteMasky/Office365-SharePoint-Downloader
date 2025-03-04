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

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PowerPointCapture:
    def __init__(self):
        self.chrome_options = Options()
        # 设置浏览器窗口大小为1920x1080
        self.chrome_options.add_argument('--window-size=1920,1080')
        self.chrome_options.add_argument('--start-maximized')
        
        # 其他必要的设置
        self.chrome_options.add_argument('--disable-infobars')
        self.chrome_options.add_argument('--disable-notifications')
        self.chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        self.chrome_options.add_experimental_option('useAutomationExtension', False)

    def find_and_click_present_button(self, driver, max_attempts=3):
        """在指定frame中查找并点击演示按钮"""
        for attempt in range(max_attempts):
            try:
                # 等待按钮出现
                present_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '[aria-label="Present"]'))
                )
                
                # 移除遮挡元素
                driver.execute_script("""
                    const overlays = document.querySelectorAll('[role="dialog"], [class*="overlay"], [class*="modal"]');
                    overlays.forEach(overlay => overlay.remove());
                """)
                
                # 确保按钮可见
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", present_button)
                time.sleep(1)
                
                # 尝试点击
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
                logging.warning(f"点击尝试 {attempt + 1} 失败: {str(e)}")
                if attempt < max_attempts - 1:
                    time.sleep(1)
                    continue
        return False

    def try_click_present_button(self, driver):
        """尝试在所有可能的frame中点击演示按钮"""
        try:
            # 首先在主页面尝试
            logging.info("在主页面中查找演示按钮...")
            if self.find_and_click_present_button(driver):
                return True
            
            # 查找所有iframe
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if not iframes:
                logging.info("未找到iframe")
                return False
            
            # 遍历每个iframe
            for i, iframe in enumerate(iframes):
                try:
                    logging.info(f"切换到iframe {i + 1}/{len(iframes)}...")
                    driver.switch_to.frame(iframe)
                    
                    if self.find_and_click_present_button(driver):
                        return True
                    
                except Exception as e:
                    logging.warning(f"处理iframe {i + 1} 时出错: {str(e)}")
                finally:
                    driver.switch_to.default_content()
            
            logging.error("在所有位置都未找到可点击的演示按钮")
            return False
            
        except Exception as e:
            logging.error(f"点击演示按钮失败: {str(e)}")
            return False

    def check_presentation_mode(self, driver):
        """检查是否已进入演示模式"""
        try:
            time.sleep(2)
            # 检查URL变化
            if "view=present" in driver.current_url.lower():
                return True
            # 检查是否全屏
            is_fullscreen = driver.execute_script("""
                return document.fullscreenElement !== null || 
                       document.webkitFullscreenElement !== null ||
                       document.mozFullScreenElement !== null ||
                       document.msFullscreenElement !== null;
            """)
            if is_fullscreen:
                return True
            # 检查特定元素
            try:
                return driver.find_element(By.CSS_SELECTOR, '[class*="presentationMode"]').is_displayed()
            except:
                return False
        except:
            return False
        
    def capture_slides(self, url, output_folder='slides'):
        driver = None
        try:
            # 创建输出文件夹
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            logging.info("启动浏览器...")
            driver = webdriver.Chrome(options=self.chrome_options)
            driver.maximize_window()
            
            logging.info("访问PowerPoint页面...")
            driver.get(url)
            time.sleep(5)  # 等待页面加载
            
            # 尝试点击演示按钮
            if not self.try_click_present_button(driver):
                # 保存页面源码以供调试
                with open('page_source.html', 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
                logging.info("页面源码已保存到 page_source.html")
                raise Exception("无法进入演示模式")
            
            # 等待演示模式加载
            time.sleep(3)
            
            # 开始截图循环
            slide_count = 0
            last_screenshot = None
            screenshots = []
            consecutive_same_count = 0  # 连续相同的次数
            
            logging.info("开始捕获幻灯片...")
            while True:
                # 等待页面加载完成
                try:
                    WebDriverWait(driver, 5).until(
                        lambda d: d.execute_script('return document.readyState') == 'complete'
                    )
                except:
                    logging.warning("等待页面加载超时")

                # 截取当前页面
                screenshot_path = os.path.join(output_folder, f'slide_{slide_count:03d}.png')
                
                # 等待动画完成
                time.sleep(2)  # 基础等待时间
                
                # 检查页面是否仍在变化
                old_source = driver.page_source
                time.sleep(0.5)
                if old_source != driver.page_source:
                    time.sleep(1.5)  # 如果页面在变化，多等待一会
                
                driver.save_screenshot(screenshot_path)
                
                # 检查是否与上一张截图相同
                if last_screenshot:
                    current_img = Image.open(screenshot_path)
                    last_img = Image.open(last_screenshot)
                    
                    if self.images_equal(current_img, last_img):
                        consecutive_same_count += 1
                        logging.info(f"检测到相同截图 ({consecutive_same_count}/10)，尝试翻到下一页...")
                        
                        if consecutive_same_count >= 10:  # 连续10次相同才确认是最后一页
                            logging.info("确认已到达最后一页")
                            os.remove(screenshot_path)
                            break
                        
                        os.remove(screenshot_path)
                        
                        # 尝试翻到下一页
                        try:
                            actions = ActionChains(driver)
                            actions.send_keys(Keys.ARROW_RIGHT)
                            actions.pause(0.5)
                            actions.perform()
                            time.sleep(2)  # 等待动画完成
                        except Exception as e:
                            logging.warning(f"发送右箭头键失败: {str(e)}")
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
                                logging.error(f"模拟按键失败: {str(e)}")
                        continue
                    else:
                        consecutive_same_count = 0
                        logging.info(f"已捕获第 {slide_count + 1} 页")
                        last_screenshot = screenshot_path
                        screenshots.append(screenshot_path)
                        slide_count += 1
                else:
                    logging.info(f"已捕获第 {slide_count + 1} 页")
                    last_screenshot = screenshot_path
                    screenshots.append(screenshot_path)
                    slide_count += 1
                
                # 模拟按右箭头键
                try:
                    actions = ActionChains(driver)
                    actions.send_keys(Keys.ARROW_RIGHT)
                    actions.pause(0.5)
                    actions.perform()
                    time.sleep(2)  # 等待动画完成
                except Exception as e:
                    logging.warning(f"发送右箭头键失败: {str(e)}")
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
                        logging.error(f"模拟按键失败: {str(e)}")
            
            # 转换为PDF
            if screenshots:
                self.create_pdf(screenshots, os.path.join(output_folder, 'presentation.pdf'))
                logging.info(f"演示文稿已保存为PDF: {os.path.join(output_folder, 'presentation.pdf')}")
            
            logging.info(f"共捕获 {len(screenshots)} 页幻灯片")
            
        except Exception as e:
            logging.error(f"发生错误: {str(e)}")
            if driver:
                # 保存错误时的页面源码
                with open('error_page_source.html', 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
                logging.info("错误页面源码已保存到 error_page_source.html")
        finally:
            if driver:
                driver.quit()
    
    def images_equal(self, img1, img2):
        """比较两张图片是否相同"""
        if img1.size != img2.size:
            return False
            
        # 转换为灰度图像并调整大小以减少计算量
        size = (img1.size[0] // 2, img1.size[1] // 2)
        img1_gray = img1.convert('L').resize(size)
        img2_gray = img2.convert('L').resize(size)
        
        # 获取图像数据
        data1 = list(img1_gray.getdata())
        data2 = list(img2_gray.getdata())
        
        # 计算差异
        total_pixels = len(data1)
        diff_pixels = sum(1 for i in range(total_pixels) if abs(data1[i] - data2[i]) > 25)
        
        # 计算差异百分比
        diff_percentage = (diff_pixels / total_pixels) * 100
        
        return diff_percentage < 2.0  # 允许2%的差异
    
    def create_pdf(self, image_files, output_file):
        """将图片合并为PDF"""
        if not image_files:
            return
            
        logging.info("正在生成PDF...")
        images = [Image.open(f) for f in image_files]
        first_image = images[0]
        first_image.save(output_file, "PDF", resolution=100.0, save_all=True, append_images=images[1:])

def main():
    try:
        url = input("请输入SharePoint PowerPoint页面的URL: ").strip()
        if not url:
            logging.error("URL不能为空")
            return
            
        output_folder = input("请输入保存文件夹名称（默认为slides）: ").strip() or 'slides'
        
        capture = PowerPointCapture()
        capture.capture_slides(url, output_folder)
        
    except KeyboardInterrupt:
        logging.info("\n程序已被用户中断")
        sys.exit(0)
    except Exception as e:
        logging.error(f"程序执行出错: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 