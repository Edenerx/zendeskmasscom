import requests
import sys
import json
import chardet
import time
import pandas as pd
import tkinter as tk
import subprocess
from tkinter import filedialog,simpledialog

group_id_choices = [20632073,4416380043282 ,20971851052178]
choices = [24403664487,536345117,7260409749138,510287123]
#assignee_id=choice
tags=[
                    "massCommunications",
                    "no_survey"
                ]
collaborator_id=24403664487
custom_fields=[
    {"id":33412807,"value":"null"}
    ]
ticket_form_id=42993
url = "https://bluesnap.zendesk.com/api/v2/tickets"
bluesnap_html_title=("""
    <style> html { background-color: #ddd; padding: 5%; } body { font-family: sans-serif; width: 90%; background-color: #fff; padding-bottom: 5%; } div { margin: 0% 2% 5% 2%; } table { width: 80%; border: 2px black solid; border-collapse: collapse; border-spacing: 10px 15px; } th, tr, td { border: 1px black solid; border-collapse: collapse; } th { padding: 5px; background-color: #163e5e; color: #fff; } td { padding: 10px; text-align: left; } ul { list-style-type: square; } ul.empty { list-style-type: none; margin: 0; padding: 0; } li { margin-bottom: 2px; } .footnote { border: none; font-size: smaller; font-style: italic; } .header { display: inline-block; margin: 0%; width: 96%; height: 4%; padding: 2%; margin-bottom: 5%; background: rgb(63, 113, 214); background: linear-gradient(90deg, rgba(63, 113, 214, 1) 60%, rgba(255, 255, 255, 1) 100%); } .logo { width: 15%; } </style> </head>
    <div class="header"> <a href="https://home.bluesnap.com/"> <img class="logo" src="https://f.hubspotusercontent10.net/hubfs/454819/Dan_Support/statusPal_2020/bluesnap-white_med.png" alt="BlueSnap Logo"> </a> </div>
    <body lang=EN-US style='tab-interval:.5in;word-wrap:break-word'>
    """)


def restart_script():
    while True:
        print("Python executable:", sys.executable)
        print("Script name:", sys.argv[0])
        print("Restarting the application...")
        print("Starting program")
        p = subprocess.Popen(['python', sys.argv[0]])
        p.wait()
        print("Program exited")
        time.sleep(5)

def choose_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        print("Selected file:", file_path)
        return file_path
    else:
        print("No file selected restarting.")
        restart_script()

def email_subject():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    while True:
        # Display the Yes/No dialog
        subject = simpledialog.askstring("Input", "Please proivde the subject", parent=root)
        if subject is None:
          restart_script()
        else:
          return subject
        
def show_list_groups():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    # Display the list as a pop-up dialog
    group_id_index = simpledialog.askinteger("Choos a group", "Select an option (enter the number):[Merchant support - 0,Business Operations - 1,Acquiring operations - 2]", 
                                             parent=root, initialvalue=0, minvalue=0, maxvalue=len(group_id_choices)-1)
    if group_id_index is None:
      restart_script()
    else:
      if group_id_index is not None:
          group_id_choice = group_id_choices[group_id_index]
          print("You chose:", group_id_choice)
          return group_id_choice
        
def show_list_users(group_id):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    if group_id == 20632073:
      # Display the list as a pop-up dialog
      index = simpledialog.askinteger("Choose", "Select an option (enter the number):[Mark - 0,Stacy - 1,Ronen - 2,Dan - 3]", parent=root, 
                                      initialvalue=0, minvalue=0, maxvalue=len(choices)-1)
      if index is None:
        restart_script()
      else:
        if index is not None:
            choice = choices[index]
            print("You chose:", choice)
            return choice
        else:
          choice = 510287123
          return choice
        
def choose_if_adding_am():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    while True:
        # Display the Yes/No dialog
        response = simpledialog.askstring("Input", "Would you like to cc AM? (yes/no) click on cancel to restart", parent=root)
        if response is None:
          restart_script()
        else:
          # Ensure the response is lowercased and only 'yes' or 'no' is returned
          if response and response.lower() in ['yes', 'no']:
              return response.lower()
          else:
              # Show an error message if the input is invalid
              tk.messagebox.showerror("Invalid Input", "Please enter 'yes' or 'no'.")

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
    
def send_email(email_requester,group_id,am_email,body_html,ticket_form_id,custom_fields,tags,email_subject_response,choice):
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
    auth=('edenm@bluesnap.com/token', 'BrzqFgjXPZiwyHqI3Eew4LizU8Q1aIDpAnOeT9t1'),
    headers=headers,
    data=payload_json 
  )
  return response

def list_of_emails(output_file,email_requester_list,group_id,list_ac_exel_path,email_subject_response,body_html,ticket_form_id,custom_fields,tags,collaborator_id):
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
     # response = simpledialog.askstring("Input", "Would you like contiune with sending the emails? (yes/no)", parent=root)
        
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

    # Output the dictionary
#      print(email_requester_list_cc)
#  else:
#      print("The required columns 'email' and 'am_email' are not found in the Excel file.")

    for email_requester in email_requester_list:
      am_email=(email_requester_list_cc[email_requester])
      print(am_email)
      if not isinstance(am_email,(int,float)):
        if email_requester in email_requester_list_cc:
        # name_cc = email_requester_list_cc[email_requester].split('@')[0]
          response=send_email(email_requester,group_id,am_email,body_html,ticket_form_id,custom_fields,tags,email_subject_response,collaborator_id)
          time.sleep(0.5)
          if  response.status_code != 201:
            print(response)
            print(json.dumps(response.json(), indent=4))
      else:
        response=send_email(email_requester,group_id,"",body_html,ticket_form_id,custom_fields,tags,email_subject_response,collaborator_id)
        time.sleep(0.5)
        if response.status_code != 201:
          print(response)
          print(json.dumps(response.json(), indent=4))
        print("Done with AM as cc")
        print(f'https://bluesnap.zendesk.com/agent/search/1?copy&type=ticket&q="{email_subject_response}"')
  else:
    for email_requester in email_requester_list:
      response=send_email(email_requester,group_id,None,body_html,ticket_form_id,custom_fields,tags,email_subject_response,collaborator_id)
      time.sleep(0.5)
      if response != [201]:
          print(response)
          print(json.dumps(response.json(), indent=4))
    print("Done without AM as cc")
    print(f'https://bluesnap.zendesk.com/agent/search/1?copy&type=ticket&q="{email_subject_response}"')

def main():
    print("""
        Please choose the user that is sending the emails.
        Then choose the excel file with the email's list
        Then choose the HTML file.
        Decide if you want to cc AM if you do then choose the excel file with the list of AM.    
        whenever you click on cancel the process will restart     
        """)
    group_id=show_list_groups()
    email_requester_list=[]
    choice=show_list_users(group_id)
    email_subject_response=email_subject()



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

    list_of_emails(list_email_exel_path, email_requester_list,group_id,list_ac_exel_path,email_subject_response,body_html,ticket_form_id,custom_fields,tags,collaborator_id)
   

#start
if __name__ == "__main__":
   main()



  
