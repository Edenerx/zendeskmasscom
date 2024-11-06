import requests
import math
import json
import chardet
import time
import pandas as pd
import tkinter as tk
from bs4 import BeautifulSoup
from tkinter import filedialog,simpledialog

def start():
   print("""
Please choose the user that is sending the emails.
Then choose the excel file with the email's list
Then choose the HTML file.
Decide if you want to cc AM if you do then choose the excel file with the list of AM.         
""")
   
def email_subject():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    while True:
        # Display the Yes/No dialog
        response = simpledialog.askstring("Input", "Please proivde the subject", parent=root)

        return response.lower()

def show_list_groups():
    
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    group_id_choices = ["groupdID0,groupdID1 ,groupdID2"]
    
    # Display the list as a pop-up dialog
    group_id_index = simpledialog.askinteger("Choos a group", "Select an option (enter the number):[Group name 0 - 0,Group name 1 - 1,Group name 2 - 2]", 
                                             parent=root, initialvalue=0, minvalue=0, maxvalue=len(group_id_choices)-1)

    if group_id_index is not None:
        group_id_choice = group_id_choices[group_id_index]
        print("You chose:", group_id_choice)
        return group_id_choice

def choose_if_adding_am():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    while True:
        # Display the Yes/No dialog
        response = simpledialog.askstring("Input", "Would you like to cc AM? (yes/no)", parent=root)
        
        # Ensure the response is lowercased and only 'yes' or 'no' is returned
        if response and response.lower() in ['yes', 'no']:
            return response.lower()
        else:
            # Show an error message if the input is invalid
            tk.messagebox.showerror("Invalid Input", "Please enter 'yes' or 'no'.")

def show_list_users(group_id):
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    if group_id == "groupdID":
      
      # Define the list of choices
      choices = ["managerID0,managerID1,managerID2,managerID3"]
      
      # Display the list as a pop-up dialog
      index = simpledialog.askinteger("Choose", "Select an option (enter the number):[name0 - 0,name1 - 1,name2 - 2,name3 - 3]", parent=root, 
                                      initialvalue=0, minvalue=0, maxvalue=len(choices)-1)

      if index is not None:
          choice = choices[index]
          print("You chose:", choice)
          return choice
    else:
       choice = "nameID"
       return choice
      
def choose_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        print("Selected file:", file_path)
        return file_path
    else:
        print("No file selected.")

def send_email(email_requester,group_id,am_email):
  payload = {
    "ticket": {
      "comment": {
        "html_body": body_html 
      },
      "priority": "low",
      "assignee_id": choice,
      "submitter_id": choice,
      "ticket_form_id": ticket_form_id,
      "custom_fields": custom_fields,
      "collaborators":am_email,
      "tags": tags,
      "group_id": group_id,
      "status": "solved",
      "subject": email_subject_response,
      "requester": email_requester,
      
    }
  }
  headers = {
    "Content-Type": "application/json",

  }

  payload_json = json.dumps(payload)



  response = requests.request(
    "POST",
    url,
    auth=('email@email.com/token', 'password'),
    headers=headers,
    data=payload_json 
    #json=payload
  )
  return response

def list_of_emails(output_file,email_requester_list,group_id,list_ac_exel_path):
  df_orig = pd.read_excel(output_file) # email 
  email_mask=df_orig['NOTIFY_EMAIL'].notna()
  mask=email_mask
  df=df_orig[mask]

  for index, row in df.iterrows():
    email_requester=row['NOTIFY_EMAIL']
    if "@" in email_requester:
      if email_requester not in email_requester_list:
          email_requester_list.append(email_requester)
      else:
        continue
    else:
      print(f"you tried to add {email_requester}, please make sure to add an email")
      response = simpledialog.askstring("Input", "Would you like contiune with sending the emails? (yes/no)", parent=root)
        
        # Ensure the response is lowercased and only 'yes' or 'no' is returned
      if response == 'no':
          break
      elif response == 'yes':
            # Show an error message if the input is invalid
          continue

  if list_ac_exel_path != None:  
    file_path = list_ac_exel_path  # Replace with your Excel file path
    df_cc = pd.read_excel(file_path)

    # Ensure the columns 'email' and 'am_email' exist in the DataFrame
    if 'NOTIFY_EMAIL' in df_cc.columns and 'am_email' in df_cc.columns:
        # Create the dictionary using 'email' as the key and 'am_email' as the value
        email_requester_list_cc = {email: am_email for email, am_email in zip(df_cc['NOTIFY_EMAIL'], df_cc['am_email'])}


    for email_requester in email_requester_list:
      am_email=(email_requester_list_cc[email_requester])
      print(am_email)
      if not isinstance(am_email,(int,float)):
        if email_requester in email_requester_list_cc:
          response=send_email(email_requester,group_id,am_email)
          time.sleep(0.5)
          if  response.status_code != 201:
            print(response)
            print(json.dumps(response.json(), indent=4))
      else:
        response=send_email(email_requester,group_id,"")
        time.sleep(0.5)
        if response.status_code != 201:
          print(response)
          print(json.dumps(response.json(), indent=4))
        print("Done with AM as cc")
        print(f'ticketsearchlink"{email_subject_response}"')
  else:
    for email_requester in email_requester_list:
      response=send_email(email_requester,group_id,None)
      time.sleep(0.5)
      if response != [201]:
          print(response)
          print(json.dumps(response.json(), indent=4))
    print("Done without AM as cc")
    print(f'ticketsearchlink"{email_subject_response}"')

def remove_double_backslashes(html_file_path):
    with open (html_file_path, "rb") as file:
       raw_data = file.read()
       result = chardet.detect(raw_data)
       encoding = result['encoding']

    with open(html_file_path, "r", encoding=encoding, errors='replace') as file:
        # Read the text file
        content =  file.read()
    
    # Remove double backslashes
    content = content.replace("\\", "")

    with open(html_file_path, "w", encoding="utf-8") as file:
        # Write the modified content back to the file
        file.write(content)
        return bluesnap_html_title + content

#start
start()

group_id=show_list_groups()
email_requester_list=[]
email_requester_list_cc={}
choice=show_list_users(group_id)

url = "URL"

bluesnap_html_title=("""
HTML TITLE
""")

email_subject_response=email_subject()
assignee_id=choice
tags=[
                "massCommunications",
                "no_survey"
            ]
collaborator_id=["collaboratorID"]
ticket_form_id="ticketformID"
custom_fields=[
{"id":"fieldID","value":"null"}
]



# Create the main window
root = tk.Tk()
root.withdraw()  # Hide the main window

# Create a pop-up dialog for choosing a file
list_email_exel_path=choose_file() # Path to your exel file
html_file = choose_file()  # Path to your HTML file
check_if_am_cc=choose_if_adding_am()
if check_if_am_cc.lower() == "yes":
  list_ac_exel_path = choose_file()
else:
   list_ac_exel_path = None
body_html=remove_double_backslashes(html_file)

# Call the function to show the list pop-up

list_of_emails(list_email_exel_path, email_requester_list,group_id,list_ac_exel_path)



  
