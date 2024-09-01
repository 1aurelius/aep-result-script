import gspread
import re
from colorama import Back, Fore, Style

def run():

    # credentials
    key_path = "SERVICE ACCOUNT KEY PATH"
    rank = "YOUR RANK"
    username = "YOUR USERNAME"

    #setup & error handler
    try:
        sa = gspread.service_account(filename=key_path)
        sh = sa.open("RandomProject")
        wks = sh.worksheet("Test")
        print(Fore.GREEN + "Connected to Google Sheets Successfully")
    except Exception as err:
        print(err)

    # finds all wave ends
    try:
        criteria = re.compile(r'WAVE\s*\d+\s*END')
        cell_list = wks.findall(criteria)
        print(Fore.GREEN + "Fetched Waves Successfully" + Fore.RESET)
    except Exception as err:
        print(err)

    # adds the row number to a list
    value_list = []
    for i in cell_list:
        value_list.append(i.row)

    # finds the max value and max index in the list
    max_value = max(value_list)
    max_index = value_list.index(max_value)

    # finds all results from the most recent wave (current wave to previous wave)
    raw_results = wks.get(f"A{value_list[max_index-1]}:X{value_list[max_index]}")

    # creation of lists
    approved_list = []
    denied_list = []

    # user input for file name
    # should be in the format of "wave-number"
    file_name = input("Enter desired file name: ")
    file = open(f"{file_name}.txt", "w")


    # iterating through raw_results and adding results to their respective lists
    length = len(raw_results) - 2 #removed 2 because of the two wave cells
    z = 0
    while z < length:
        z += 1
        if (raw_results[z][19] == 'Denied'):
            denied_list.append(raw_results[z])
        if (raw_results[z][19] == 'Approved'):
            approved_list.append(raw_results[z])

    # reading off denied results
    x = 0
    # checking if the length of denied list isn't 0 so i don't get ugly lines
    if (len(denied_list) != 0):
        file.write("[ DENIED ]\n\n")
        while x < len(denied_list):
            file.write(f"Greetings, {denied_list[x][1]}. \nI am {rank} {username}, from the AEP staff, contacting you regarding your AEP application results. \nThe AEP staff is sorry to inform you that you have **failed** the application. \n\nThis program is tough to get into and progress in so failing is totally understandable. Do not be discouraged in trying again, we always welcome and notice dedication. Your notes on what to improve on are listed below: \n\n{denied_list[x][21]} before trying again.\n\n\n")
            file.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\n\n")
            x += 1

    # reading off approved results
    y = 0
    # checking if the length of approved list isn't 0 so i don't get ugly lines
    if (len(approved_list) != 0):
        file.write("[ APPROVED ]\n\n")
        while y < len(approved_list):
            file.write(f"Greetings, {approved_list[y][1]}. \nI am {rank} {username} from the AEP staff, contacting you regarding your AEP application results. \nThe AEP staff is very glad to inform you that you have been **accepted** into the program. \n\nThe Cerberus Alternative Entrance Program functions as an alternative entrance into Cerberus that differs from Observational Tryouts. The program is designed to find and pick out those that are worthy of being a Cerberus Operative without the constrictions of timezones and OTs. \n\nUpon entry into the AEP you will be given restricted access into official Cerberus channels and VCs. All information within these channels and VCs will remain confidential. \n\nBiweekly voting sessions are held by the Cerberus AEP Staff to determine who is prepared to progress to the Novice rank. In order to progress, you must show us competence performing the duties of a Cerberus member in this time frame. \n\nOnce again welcome onboard the AEP program, any questions can be forwarded to me. \n\nGuidelines: https://docs.google.com/document/d/1CevHmFO7UP5QwIf14lJ4ncdnYRK01NwNCU_ak0Nzt_A/edit\n\n\n")
            file.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\n\n")
            y += 1

run()
