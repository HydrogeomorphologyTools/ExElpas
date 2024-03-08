from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
import openpyxl

def login(driver):
    driver.get("https://epas.amministrazione.cnr.it/oauth/login")

    username = driver.find_element(By.ID, "username")
    password = driver.find_element(By.ID, "password")

    username.send_keys("cnr_epas_USERNAME_HERE")  # safer alternative is using input for entering credentials interactively
    password.send_keys("cnr_epas_PASSWORD_HERE")

    driver.find_element(By.NAME, "login").click()

def navigate_to_month(driver, month, year, person_id):
    # Construct the URL with the desired month, year, and personId
    url = f"https://epas.amministrazione.cnr.it/stampings/stampings?month={month}&year={year}&personId={person_id}&officeId=92&day=5"
    driver.get(url)

def retrieve_tempo_lavoro_values(driver):
    # Wait for the main table to load
    main_table = driver.find_element(By.ID,
                                     "tabellonetimbrature")  # "tabellonetimbrature" = main table id
    rows = main_table.find_elements(By.TAG_NAME, "tr")[1:]  # Start from the second row (skip header go to first day)

    data = []

    # Loop through each row in the table
    for row in rows:
        # Initialize festivi, capitalized, and assenza values as empty strings
        festivi_value = ""
        capitalized_value = ""
        assenza_value = ""
        altroferiemalattia = "00:00"  # Default value for altroferiemalattia

        # Try to find festivi, capitalized, or assenza cell in the row
        try:
            festivi_cell = row.find_element(By.CSS_SELECTOR, "td.festivi.default-single")
            festivi_value = "F_" + festivi_cell.text.strip()
        except:
            pass

        try:
            capitalized_cell = row.find_element(By.CSS_SELECTOR, "td.capitalized.default-single")
            capitalized_value = capitalized_cell.text.strip()
        except:
            pass

        try:
            assenza_cell = row.find_element(By.CSS_SELECTOR, "td.assenza.default-single")
            assenza_value = assenza_cell.text.strip()

            # Check if any of the specified codes is a substring of assenza_value
            specified_codes = ["91", "31", "32", "37", "21P"]
            if any(code in assenza_value for code in specified_codes):
                altroferiemalattia = "07:12"
        except:
            pass

        # Find the tempo lavoro cell in the row
        tempo_lavoro_cell = row.find_element(By.CSS_SELECTOR, "td.tempoLavoro.default-single")

        # Get the text value from the tempo lavoro cell
        tempo_lavoro_value = tempo_lavoro_cell.text.strip()

        # Append the data to the list
        data.append([festivi_value or capitalized_value, tempo_lavoro_value, assenza_value, altroferiemalattia])

        # Print the festivi, capitalized, assenza, and tempo lavoro values for the current row
        print("Giorno:", festivi_value or capitalized_value, "TempoOreProduttive:", tempo_lavoro_value, "Cod_Assenza:", assenza_value, "AltroFerieMalattia:", altroferiemalattia)

    return data

# Configure Selenium options
options = Options()
options.add_argument("--headless")
driver = webdriver.Edge(options=options)

# Log in to the page
login(driver)

# Choose the desired month, year, personId and filename prefix
month = 2
year = 2024
person_id = 0000  # e.g., (optional) 4 digits of id personal code available in authenticated epas url
prefix_xls = "Pippo"  # filename_prefix for Excel saveout file

# Navigate to the desired month, year, and personId
navigate_to_month(driver, month, year, person_id)

# Retrieve festivi or capitalized, assenza, and tempo lavoro values
data = retrieve_tempo_lavoro_values(driver)

# Close the browser
driver.quit()

# Transpose the data
transposed_data = list(map(list, zip(*data)))

# Create a new Excel workbook
wb = openpyxl.Workbook()
sheet = wb.active

# Add headers
headers = ["Giorno", "TempoOreProduttive", "Cod_Assenza", "AltroFerieMalattia"]
for header in headers:
    sheet.append([header])

# Add transposed data starting at B2
for i, row in enumerate(transposed_data):
    for j, value in enumerate(row):
        sheet.cell(row=i+1, column=j+2, value=value)

# Save the workbook
filename = f"{prefix_xls}_epas_{year}_{month}.xlsx"
wb.save(filename)
