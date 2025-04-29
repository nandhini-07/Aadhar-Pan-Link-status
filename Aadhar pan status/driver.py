from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# Set Chrome options (optional, but recommended for headless or custom configurations)
chrome_options = Options()
chrome_options.add_argument("--start-maximized")  # Optional: To start Chrome maximized

# Use ChromeDriverManager to install the correct driver version
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

