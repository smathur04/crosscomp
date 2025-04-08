import os
import sys
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchWindowException


start = input("Hit enter if you would like to skip sim check, otherwise type 'yes' and then click enter: ")


driver = webdriver.Chrome()
driver.get("https://beta.rap.t-mobile.com/rap/home")


try:
    script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    os.chdir(script_dir) 
    file_path = [f for f in os.listdir() if f.endswith(".xlsx")][0]
    xls = pd.ExcelFile(file_path)
    df = xls.parse(sheet_name=xls.sheet_names[1], header=None)
    logins_list = df.iloc[0:5, 0].tolist()
    if len(logins_list) == 4:
        logins_list.append("default")
    elif len(logins_list) == 5 and str(logins_list[4]).strip() == "":
        logins_list[4] = "default"
    df = xls.parse(sheet_name=xls.sheet_names[0])
    df = df.astype(str)
    sim_col = next(col for col in df.columns if "sim" in col.lower())
    phone_col = next(col for col in df.columns if "phone" in col.lower())
    sim_dict = {sim: "" if pd.isna(df.at[i, phone_col]) or df.at[i, phone_col] == "" else df.at[i, phone_col] for i, sim in df[sim_col].items()}
    sim_dict_final = {index + 1: {key: value} for index, (key, value) in enumerate(sim_dict.items())}
except Exception as e:
    raise Exception(f"There is something wrong with your excel here is the error the program produced \n {e}")


try:
    email_input = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "i0116")))
    email_input.send_keys(logins_list[0])
    email_input.send_keys(Keys.RETURN)
    pass_input = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "i0118")))
    pass_input.send_keys(logins_list[1])
    pass_input.send_keys(Keys.RETURN)
except Exception as e:
    raise Exception(f"There is something wrong with your tmobile login here is the error the program produced \n {e}")


try:
    zip_one = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-118")))
    street = driver.find_element(By.ID, "rap-store-street").text
    zip = driver.find_element(By.ID, "rap-store-city-zip").text[-5:]
    if logins_list[4] == "default": logins_list[4] = zip
    pin = driver.find_element(By.ID, "rap-dealer-code").text
    imei = 356656426318563
    zip_one.clear()
    zip_one.send_keys(zip)
    zip_one.send_keys(Keys.RETURN)
except Exception as e:
    raise Exception(f"There is something wrong at the start here is the error the program produced \n {e}")


def sim_check():
    imei_send = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, f"tmo-input-default-74")))
    imei_send.send_keys(imei)
    check_compat = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "checkCompatibility")))
    try:
        driver.execute_script("arguments[0].click();", check_compat)
    except:
        check_compat.click()

    for row_number, row_sim_dict in sim_dict_final.items():
        for row_sim in row_sim_dict.keys():
            sim = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-76")))
            sim.clear()
            sim.send_keys(row_sim)
            valsim = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "checkSimValidationButton")))
            try:
                driver.execute_script("arguments[0].click();", valsim)
            except:
                valsim.click()
            try:
                error_element = WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.ID, "errorMessage0")))
                error_text = error_element.text
                sim_dict_final[row_number][row_sim] = error_text
                df.at[row_number - 1, phone_col] = sim_dict_final[row_number][row_sim]
                with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                    df.to_excel(writer, sheet_name=xls.sheet_names[0], index=False)
            except:
                go_back = WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.XPATH, "//h4[contains(text(), 'Enter your T-Mobile SIM card number')]")))
                try:
                    driver.execute_script("arguments[0].click();", go_back)
                except:
                    go_back.click()
                continue
    driver.refresh()    


def clean_up():
    for i in range((len(sim_dict_final.keys())) - (len(sim_dict_final.keys()) % 5) + 1, (len(sim_dict_final.keys())) + 1):
        del sim_dict_final[i]

    have_to_go = set()

    row_number = 1
    while row_number <= len(sim_dict_final):
        row_phone = sim_dict_final[row_number].values()
        row_phone = list(row_phone)[0]
        if row_phone.strip() not in ["", "nan"] :
            if row_number > 5:
                base = ((row_number - 1) // 5) * 5 + 1
                have_to_go.add(base)
                have_to_go.add(base + 1)
                have_to_go.add(base + 2)
                have_to_go.add(base + 3)
                have_to_go.add(base + 4)
                row_number = base + 5
            else:
                base = 5
                have_to_go.add(base)
                have_to_go.add(base - 1)
                have_to_go.add(base - 2)
                have_to_go.add(base - 3)
                have_to_go.add(base - 4)
                row_number = 6
        else:
            row_number += 1

    for j in have_to_go:
        del sim_dict_final[j]


def line(current_row, current_sim, x):
    imei_send = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, f"tmo-input-default-74")))
    imei_send.send_keys(imei)
    check_compat = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "checkCompatibility")))
    try:
        driver.execute_script("arguments[0].click();", check_compat)
    except:
        check_compat.click()

    sim = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-76")))
    sim.clear()
    sim.send_keys(current_sim)

    valsim = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "checkSimValidationButton")))
    try:
        driver.execute_script("arguments[0].click();", valsim)
    except:
        valsim.click()

    if x != 21:
        switch_prepaid = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, f"ui-tabpanel-{x}-label")))
        try:
            driver.execute_script("arguments[0].click();", switch_prepaid)
        except:
            switch_prepaid.click()
        ninesev = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='TPP100TTUNLw10XPROMO-4']/parent::span/parent::label")))
        try:
            driver.execute_script("arguments[0].click();", ninesev)
        except:
            ninesev.click()
    else: 
        twofi = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='CbTM25_TTUw8Xc_1mo-2']/parent::span/parent::label")))
        try:
            driver.execute_script("arguments[0].click();", twofi)
        except:
            twofi.click()
    time.sleep(4)
    confour = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "continueToStepFourButton")))
    try:
        driver.execute_script("arguments[0].click();", confour)
    except:
        confour.click()

    zip_two = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "cta-default-116")))
    zip_two.clear()
    zip_two.send_keys(logins_list[4])
    search_icon = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//a[@class='hover-magenta search-icon']")))
    try:
        driver.execute_script("arguments[0].click();", search_icon)
    except:
        search_icon.click()
    time.sleep(4)

    emergopt = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='tmo-radio-button-opt-in-emergency-0']/parent::span/parent::label")))
    try:
        driver.execute_script("arguments[0].click();", emergopt)
    except:
        emergopt.click()
    time.sleep(1)
    try:
        streetaddy = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "line-setup-e911-addressLine1-input-0")))
        streetaddy.clear()
        streetaddy.send_keys(street)
    except:
        streetaddy = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "line-setup-e911-addressLine1-input-0")))
        streetaddy.clear()
        streetaddy.send_keys(street)
    zip_three = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "line-setup-e911-postalCode-input-0")))
    zip_three.clear()
    zip_three.send_keys(zip) 
    time.sleep(1)
    try:
        streetaddy.click()
    except:
        driver.execute_script("arguments[0].click();", streetaddy)
    time.sleep(5)
    suggest_add = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//input[@id='tmo-radio-button-form-input-value-3_1']/parent::span/parent::label")))
    try:
        driver.execute_script("arguments[0].click();", suggest_add)
    except:
        suggest_add.click()
    add_continue = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "address-validator-select-button")))
    try:
        driver.execute_script("arguments[0].click();", add_continue)
    except:
        add_continue.click()
    complete_line_setup = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "button-on-complete-line-setup")))
    try:
        driver.execute_script("arguments[0].click();", complete_line_setup)
    except:
        complete_line_setup.click()

    phone = driver.find_element(By.XPATH, "//p[contains(text(), 'Phone number: ')]")
    sim_dict_final[current_row][current_sim] = phone.text.replace("Phone number: ", "").strip()
    print(sim_dict_final[current_row][current_sim])
    df.at[current_row - 1, phone_col] = sim_dict_final[current_row][current_sim]
    with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=xls.sheet_names[0], index=False)

    if x != 21:
        nexty = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "navigate-on-continuebtn")))
        try:
            driver.execute_script("arguments[0].click();", nexty)
        except:
            nexty.click()


def checkout():
    wowcont = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "button-on-continue")))
    try:
        driver.execute_script("arguments[0].click();", wowcont)
    except:
        wowcont.click()

    emails = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-92")))
    emails.send_keys(email) 
    confemail = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-93")))
    confemail.send_keys(email) 
    pins = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-pin-94")))
    pins.send_keys(pin) 
    confpin = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-pin-95")))
    confpin.send_keys(pin) 
    time.sleep(1)
    pins.click()

    wowconttwo = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "backToTop")))
    try:
        driver.execute_script("arguments[0].click();", wowconttwo)
    except:
        wowconttwo.click()

    epay_email = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "inputEmail")))
    epay_email.send_keys(logins_list[2]) 
    epay_pass = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "inputPassword")))
    epay_pass.send_keys(logins_list[3]) 

    sign_in_button = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.XPATH, "//button[span[normalize-space(text())='Sign In']]")))
    try:
        driver.execute_script("arguments[0].click();", sign_in_button)
    except:
        sign_in_button.click()

    wowcontthree = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "submitTransaction")))
    try:
        driver.execute_script("arguments[0].click();", wowcontthree)
    except:
        wowcontthree.click()

    wowcontfour = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "transactionDone")))
    try:
        driver.execute_script("arguments[0].click();", wowcontfour)
    except:
        wowcontfour.click()

    tostart = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "go-to-account-button")))
    try:
        driver.execute_script("arguments[0].click();", tostart)
    except:
        tostart.click()

    zip_one = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-118")))
    zip_one.clear()
    zip_one.send_keys(zip)
    zip_one.send_keys(Keys.RETURN)


def run_bot(check):
    if len(sim_dict_final.keys()) < 5:
        raise Exception("You need at least 5 sims bro!")
    if check: 
        clean_up()
        sim_check()
        zip_one = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-118")))
        zip_one.clear()
        zip_one.send_keys(zip)
        zip_one.send_keys(Keys.RETURN)
    else:
        clean_up()
    cut_off = 1
    save = set()
    for row_number, row_sim_dict in sim_dict_final.items():
        try:
            for row_sim in row_sim_dict.keys():
                if row_number in save:
                    continue
                global email
                email = f"s{row_sim[-7:]}@gmail.com"
                cut_off = row_number
                if row_number % 5 == 1:
                    line(row_number, row_sim, 1)
                elif row_number % 5 == 2:
                    line(row_number, row_sim, 7)
                elif row_number % 5 == 3:
                    line(row_number, row_sim, 12)
                elif row_number % 5 == 4:
                    line(row_number, row_sim, 17)
                else:
                    line(row_number, row_sim, 21)
                    checkout()
        except NoSuchWindowException as e:
            return
        except Exception as e:
            print(f"Something went wrong when doing SIM:{sim_dict_final[cut_off]}. Here is the error the program produced \n {e}")
            temp = set()
            if cut_off > 5:
                base = ((row_number - 1) // 5) * 5 + 1
                temp.add(base)
                temp.add(base + 1)
                temp.add(base + 2)
                temp.add(base + 3)
                temp.add(base + 4)
            else:
                base = 5
                temp.add(base)
                temp.add(base - 1)
                temp.add(base - 2)
                temp.add(base - 3)
                temp.add(base - 4)
            for j in temp:
                df.at[j - 1, phone_col] = "*MANUAL CHECK NEEDED*" + str(df.at[j - 1, phone_col])
                with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                    df.to_excel(writer, sheet_name=xls.sheet_names[0], index=False)
            save.update(temp)
            driver.get("https://beta.rap.t-mobile.com/rap/home")
            time.sleep(2)
            zip_one = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "tmo-input-default-118")))
            zip_one.clear()
            zip_one.send_keys(zip)
            zip_one.send_keys(Keys.RETURN) 
            continue
    print("ADDING EMAILS TO EXCEL...")
    df["EMAIL"] = ""
    for i in range(0, len(df), 5):
        if i + 4 < len(df):
            last_7 = df.at[i + 4, sim_col][-7:]  
            epic_email = f"s{last_7}@gmail.com"
            df.loc[i:i+4, "EMAIL"] = [epic_email] + [""] * 4 
    with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=xls.sheet_names[0], index=False)
    print("BOT DONE!")


run_bot(check = bool(start))