

### message sent to each customer ######## 
# Scroll to the bottom to run this file


#import os
#os.getcwd()
#os.chdir(r"C:\Users\Nitish\Desktop\g work\sending mails to user")

def message(name, username, salutation, mobile) : 
    import random
    keyword_string = "qwertyuioplkjhgfdsazxcvbnmQAWSEDRFGTYHJUIKLOPZXCVBNM0123456789~`!@#$%^&*()_+-*|}]{[:;'?>.<,"
    password = "".join(random.sample(keyword_string, 8))
    
    website_link = "www.abcCompany.com/change.password/"

    string_output = f"""
    Dear {name}, \n
    Thank you for registering on abcCompany.
    We are glad to inform that your account is now activated.
    We will send updates on your registered mobile number {mobile}.
    Also {salutation}{name}, you can send us queries anytime.
    Please save your username and password:-
        username = {username}
        password = {password}
          
    Kindly visit below link to change your password:
    {website_link}
    """
    return string_output


def write_to_file(r,string_output, fileNameStart) : 
    filePath = r"C:\Users\Nitish\Desktop\g work\sending mails to user\message sent folder"
    fileName = f"\{r+1}. {fileNameStart}_email_message.txt"
    fileHandler = open((filePath + fileName), "w")
    fileHandler.write(string_output)
    fileHandler.close()


def excel_table() : 

    import pandas as pd
    import numpy as np
    
    tableName = pd.read_excel("customer.xlsx")
#    print(tableName.head())
    
    tableName['Name'] = tableName['First Name'] + " " + tableName['Last Name']
    
#    print(tableName.columns)
    
#    tableName = tableName[['Name', "Email", "Mobile", "Gender"]]
#    print(tableName.head())
    
    rows = tableName.shape[0]
    
    for r in range(rows) : 
        name = tableName.loc[r,"Name"]
        username = tableName.loc[r,"Email"]
        gender = tableName.loc[r,"Gender"]
        mobile = tableName.loc[r,"Mobile"]
        fileNameStart = tableName.loc[r, "First Name"]
        
    #    print(f'{name}, {username}, {gender}, {mobile}')
        if gender == "Male" : 
            salutation = "Mr."
        else : 
            salutation = "Ms."
            
        string_output = message(name, username, salutation, mobile)
        write_to_file(r,string_output, fileNameStart)
    
    
    
import time
start = time.perf_counter()

## run below for whole  program
excel_table()

end = time.perf_counter()

print("Time taken to complete this program = ", int(end - start), " seconds.")





















