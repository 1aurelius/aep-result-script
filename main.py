import gspread
import re
from colorama import Back, Fore, Style

key_path = "PATH TO SERVICE ACCOUNT KEY" # <-- CHANGE THIS

# google sheets api
try:
    sa = gspread.service_account(filename=key_path)
    sh = sa.open("RandomProject")
    wks = sh.worksheet("Test")
    print(Fore.GREEN + "\nConnected to Google Sheets Successfully" + Fore.RESET)
except gspread.exceptions.SpreadsheetNotFound:
    print(Fore.RED + "A valid spreadsheet was unable to be found, try double checking the spreadsheet name." + Fore.RESET)
except gspread.exceptions.WorksheetNotFound:
    print(Fore.RED + "A valid worksheet was unable to be found, try double checking the spreadsheet name." + Fore.RESET)
except Exception as err:
    print(err)

def get_wave(file_name):
    # credentials
    rank = "YOUR RANK"
    username = "YOUR USERNAME"

    # finds all wave ends
    try:
        criteria = re.compile(r'WAVE\s*\d+\s*END')
        wave_results = wks.findall(criteria)
        print(Fore.GREEN + "\nFetched Waves Successfully" + Fore.RESET)
    except Exception as err:
        print(err)

    # adds the row number to a list
    value_list = [i.row for i in wave_results]
    cleaned_results = wks.get(f"A{value_list[-1]}:X{value_list[-2]}") # these results still have the wave end cells in them

    # user input for file name
    # should be in the format of "wave-number"
    file = open(f"{file_name}.txt", "w")

    # list creation
    approved_list = []
    denied_list = []

    # iterating through raw_results and adding results to their respective lists
    length = len(cleaned_results) - 2 #removed 2 because of the two wave cells
    z = 0
    while z < length:
        z += 1
        if cleaned_results[z][19] == 'Denied':
            denied_list.append(cleaned_results[z])
        if cleaned_results[z][19] == 'Approved':
            approved_list.append(cleaned_results[z])

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
        file.write("\n\n\n[ APPROVED ]\n\n")
        while y < len(approved_list):
            file.write(f"Greetings, {approved_list[y][1]}. \nI am {rank} {username} from the AEP staff, contacting you regarding your AEP application results. \nThe AEP staff is very glad to inform you that you have been **accepted** into the program. \n\nThe Cerberus Alternative Entrance Program functions as an alternative entrance into Cerberus that differs from Observational Tryouts. The program is designed to find and pick out those that are worthy of being a Cerberus Operative without the constrictions of timezones and OTs. \n\nUpon entry into the AEP you will be given restricted access into official Cerberus channels and VCs. All information within these channels and VCs will remain confidential. \n\nBiweekly voting sessions are held by the Cerberus AEP Staff to determine who is prepared to progress to the Novice rank. In order to progress, you must show us competence performing the duties of a Cerberus member in this time frame. \n\nOnce again welcome onboard the AEP program, any questions can be forwarded to me. \n\nGuidelines: https://docs.google.com/document/d/1CevHmFO7UP5QwIf14lJ4ncdnYRK01NwNCU_ak0Nzt_A/edit\n\n\n")
            file.write("@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\n\n")
            y += 1

def get_user_results(user):
    try:
        user_results = wks.findall(user)
        user_list = [i.row for i in user_results]
        answer_list = wks.get(f"G{user_list[-1]}:S{user_list[-1]}")
        answer_list = answer_list[0] # "un-nesting" the nested list from the wks.get above
    except Exception as err:
        print(err)

    question_list = [
    '**Do you prefer working individually or with peers and why?**',
    '**How would you benefit the subdivision if you manage to pass the program?**',
    '**Have you had any experiences as an elite, if so what were they?**',
    '**What do you want to learn from the program?**',
    '**Where are Cerberus on the TNI Hierarchy?**',
    '**What would you do if you saw an abusive Cerberus on the border?**',
    "**Which of the listed ranks can allow Cerberus Operatives to raid? (Assume they're only in Cerberus)**",
    '**When given the authority, describe all situations where you would suspend mandatory positions.**',
    '**While patrolling as a Cerberus, a citizen is constant following and trying to entice you. Would tasing them to stop being a nuisance be reasonable?**',
    '**Are you willing to be an active member of the program and try your best to progress?**',
    "**Do you agree that all answers you have provided are your own words and not anybody else's?**",
    '**Will you adhere to all future rules and expectations placed on you if you are accepted into the AEP?**']

    print(f"""
{question_list[0]}
> {answer_list[0]}

{question_list[1]}
> {answer_list[1]}

{question_list[2]}
> {answer_list[2]}

{question_list[3]}
> {answer_list[3]}

{question_list[4]}
> {answer_list[4]}

{question_list[5]}
> {answer_list[5]}

{question_list[6]}
> {answer_list[6]}

{question_list[7]}
> {answer_list[7]}

{question_list[8]}
> {answer_list[9]}

{question_list[9]}
> {answer_list[10]}

{question_list[10]}
> {answer_list[11]}

{question_list[11]}
> {answer_list[12]}
    """)

def run():
    options = input("""
1. Get the current wave results
2. Get the most recent answers for an applicant
                        
Make a selection: """)

    if options == '1':
        get_wave(input("\nDesired file name: "))
    if options == '2':
        requested_user = input("\nRequested Username: ")
        get_user_results(requested_user)
    else:
        print("err")

run()
