from tkinter import Tk, Label, Entry, Button, StringVar, font, Text
from tkinter.scrolledtext import ScrolledText
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
from io import StringIO
import sys
import threading

# Redirect console output to a StringIO object
console_output = StringIO()
sys.stdout = console_output

def login(driver):
    username = username_var.get()
    password = password_var.get()

    driver.get("https://epas.amministrazione.cnr.it/oauth/login")

    username_input = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    password_input = WebDriverWait(driver, 2).until(
        EC.presence_of_element_located((By.ID, "password"))
    )

    username_input.send_keys(username)
    password_input.send_keys(password)

    driver.find_element(By.NAME, "login").click()

# craft the epas url
def navigate_to_month(driver, month, year, person_id):
    url = f"https://epas.amministrazione.cnr.it/stampings/stampings?month={month}&year={year}&personId={person_id}&officeId=92&day=5"
    driver.get(url)

def retrieve_tempo_lavoro_values(driver):
    try:
        main_table = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.ID, "tabellonetimbrature"))
        )
    except Exception as e:
        console_output.write(f"Error: {e}\n")
        return []

    rows = main_table.find_elements(By.TAG_NAME, "tr")[1:]

    data = []

    for row in rows:
        festivi_value = ""
        capitalized_value = ""
        assenza_value = ""
        altroferiemalattia = "00:00"  # Default value for altroferiemalattia

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
            specified_codes = ["91", "31", "32", "37", "21P"]  # to be continued (integrated)
            if any(code in assenza_value for code in specified_codes):
                altroferiemalattia = "07:12"
        except:
            pass

        tempo_lavoro_cell = row.find_element(By.CSS_SELECTOR, "td.tempoLavoro.default-single")
        tempo_lavoro_value = tempo_lavoro_cell.text.strip()

        # Append the data to the list
        data.append([festivi_value or capitalized_value, tempo_lavoro_value, assenza_value, altroferiemalattia])

        # Print the festivi, capitalized, assenza, and tempo lavoro values for the current row
        console_output.write(f"Giorno: {festivi_value or capitalized_value} TempoOreProduttive: {tempo_lavoro_value} CodAssenza: {assenza_value} AltroFerieMalattia: {altroferiemalattia}\n")

    return data

def retrieve_data_and_save():
    month = month_var.get()
    year = year_var.get()
    person_id = person_id_var.get()
    prefix_xls = prefix_xls_var.get()

    options = Options()
    options.add_argument("--headless")
    driver = webdriver.Edge(options=options)

    def update_console_output():
        console_text = console_output.getvalue()
        console_text_widget.config(state="normal")
        console_text_widget.delete(1.0, "end")
        console_text_widget.insert("end", console_text)
        console_text_widget.config(state="disabled")
        console_text_widget.see("end")
        console_text_widget.update_idletasks()

    def retrieve_data():
        update_console_output()
        console_output.write("\n")
        console_output.write("Retrieving data...\n")

        login(driver)
        navigate_to_month(driver, month, year, person_id)
        console_output.write("Navigating to the specified month and year...\n")

        data = retrieve_tempo_lavoro_values(driver)
        console_output.write("Data retrieved successfully.\n")
        driver.quit()

        transposed_data = list(map(list, zip(*data)))

        wb = openpyxl.Workbook()
        sheet = wb.active
        headers = ["Giorno", "TempoOreProduttive", "CodAssenza", "AltroFerieMalattia"]
        for header in headers:
            sheet.append([header])

        for i, row in enumerate(transposed_data):
            for j, value in enumerate(row):
                sheet.cell(row=i+1, column=j+2, value=value)

        filename = f"{prefix_xls}_epas_{year}_{month}.xlsx"
        wb.save(filename)

        console_output.write(f"Data saved to {filename}.\n")
        update_console_output()

    # Run the retrieval function in a separate thread
    retrieval_thread = threading.Thread(target=retrieve_data)
    retrieval_thread.start()

# Initialize the Tkinter window
window = Tk()
window.title("EPAS Data Retrieval")

# Define StringVar variables to hold the user input
username_var = StringVar()
password_var = StringVar()
month_var = StringVar()
year_var = StringVar()
person_id_var = StringVar()
prefix_xls_var = StringVar()

# Set font size
font_size = font.Font(size=14)

# Define labels and entry widgets for user input
Label(window, text="Username:", font=font_size).grid(row=0, column=0, sticky="E", padx=10, pady=5)
Entry(window, textvariable=username_var, font=font_size).grid(row=0, column=1, padx=10, pady=5)
Label(window, text="Password:", font=font_size).grid(row=1, column=0, sticky="E", padx=10, pady=5)
Entry(window, textvariable=password_var, show="*", font=font_size).grid(row=1, column=1, padx=10, pady=5)
Label(window, text="Month (1-12):", font=font_size).grid(row=2, column=0, sticky="E", padx=10, pady=5)
Entry(window, textvariable=month_var, font=font_size).grid(row=2, column=1, padx=10, pady=5)
Label(window, text="Year:", font=font_size).grid(row=3, column=0, sticky="E", padx=10, pady=5)
Entry(window, textvariable=year_var, font=font_size).grid(row=3, column=1, padx=10, pady=5)
Label(window, text="Person ID:", font=font_size).grid(row=4, column=0, sticky="E", padx=10, pady=5)
Entry(window, textvariable=person_id_var, font=font_size).grid(row=4, column=1, padx=10, pady=5)
Label(window, text="Prefix for Filename:", font=font_size).grid(row=5, column=0, sticky="E", padx=10, pady=5)
Entry(window, textvariable=prefix_xls_var, font=font_size).grid(row=5, column=1, padx=10, pady=5)

# Define button to start data retrieval
Button(window, text="Retrieve Data", command=retrieve_data_and_save, font=font_size).grid(row=6, columnspan=2, pady=10)

# Define a Text widget to display console output
console_text_widget = ScrolledText(window, width=40, height=10, font=font_size)
console_text_widget.grid(row=0, column=2, rowspan=7, padx=10, pady=5, sticky="NSEW")
console_text_widget.config(state="disabled")

# Run the Tkinter event loop
window.mainloop()
