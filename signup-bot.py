import pandas as pd
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import threading

def run_script():
    file_path = selected_file_label["text"]
    list_taken_emails = []

    df = pd.read_excel(file_path)

 

    def update_progress_bar(i):
        progress_bar["value"] = int((i+1)/df.shape[0] * 100)

    def run_loop():
        for i in range(df.shape[0]):
            chrome_path = r"C:\Selenium\chromedriver.exe"
            driver = webdriver.Chrome(chrome_path)
            email = df['emails'][i]

            url = r"https://yourlink.com/"
            driver.get(url)

            password = driver.find_element(By.ID, "member_password")
            password.clear()

            password.send_keys(email)

            repassword = driver.find_element(By.ID, "member_password_confirmation")
            repassword.clear()

            repassword.send_keys(email)

            email_input = driver.find_element(By.ID, "member_email")
            email_input.clear()

            email_input.send_keys(email)

            loginBtn = driver.find_element(By.CLASS_NAME, "memberaccessButton")
            elements = driver.find_elements(By.CLASS_NAME, "close")
            loginBtn.click()
            try:
                element = driver.find_element(By.XPATH, "//*[contains(text(), 'Email has already been taken')]")
                print("Email has already been taken")
                list_taken_emails.append(email)
            except:
                print("Email is available")

            print("No. of Signups so far: {}".format(i+1), email)

            root.after(10, update_progress_bar, i)

        driver.quit()

        with xlsxwriter.Workbook('taken_emails.xlsx') as workbook:
            worksheet = workbook.add_worksheet()
            for row_num, email in enumerate(list_taken_emails):
                worksheet.write(row_num, 0, email)

        print("Email has already been taken",list_taken_emails)

    t = threading.Thread(target=run_loop)
    t.start()

root = tk.Tk()
root.title("Signup bot")
root.iconbitmap("**.ico")
root.geometry("380x160")
root.configure(bg="#66347F")  

selected_file_label = tk.Label(root, text="")
selected_file_label.pack()

def select_file_and_run():
    file_path = filedialog.askopenfilename(filetypes=[('Excel file ', '*.xlsx')])
    selected_file_label.config(text=file_path)
    run_script_button.config(state="normal")

select_file_button = tk.Button(root, text="Select Email Excel File", command=select_file_and_run)
select_file_button.pack()

run_script_button = tk.Button(root, text="Start", command=run_script, state="disabled")
run_script_button.pack()

progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=200, mode="determinate")
progress_bar.pack()

root.mainloop()