import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By

# 1. Настраиваем браузер так, чтобы он выглядел МАКСИМАЛЬНО как человеческий
options = webdriver.ChromeOptions()

# ВАЖНО: Мы убрали строку options.add_argument("--headless")
# Теперь вы увидите, как откроется настоящее окно браузера.

# Отключаем флаг автоматизации, который видят защиты сайтов
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

# Запускаем браузер
driver = webdriver.Chrome(options=options)

try:
    url = "https://quotes.toscrape.com/"
    print("Открываю браузер...")
    driver.get(url)
    
    # Даем сайту 5 секунд, чтобы он полностью загрузил весь контент
    print("Ожидаю загрузку страницы...")
    time.sleep(5)
    
    # Ищем карточки
    quote_blocks = driver.find_elements(By.CLASS_NAME, "quote")
    
    if not quote_blocks:
        print("Элементы не найдены. Выведите в терминал код страницы для проверки:")
        # Если пусто, выведем первые 500 символов страницы, чтобы понять, что там
        print(driver.page_source[:500]) 
    else:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Цитаты"
        sheet.append(["Цитата", "Автор"])
        
        for block in quote_blocks:
            text = block.find_element(By.CLASS_NAME, "text").text
            author = block.find_element(By.CLASS_NAME, "author").text
            sheet.append([text, author])
            
        wb.save("quotes_selenium.xlsx")
        print(f"\n[УСПЕХ] Собрано цитат: {len(quote_blocks)}.")
        print("Файл quotes_selenium.xlsx успешно создан в папке с проектом!")

except Exception as e:
    print(f"Произошла ошибка во время выполнения: {e}")

finally:
    # Оставляем браузер открытым на 3 секунды, чтобы вы успели увидеть результат, и закрываем
    time.sleep(3)
    driver.quit()
