import os, random, clipboard, threading, pythoncom, logging, getpass, requests, json, re, time
import tkinter as tk
from tkinter import simpledialog
import tkinter.font as tkfont
import tkinter.messagebox as messageBox
from pyad import adquery, aduser, adgroup, adcontainer, pyadutils
from PIL import Image, ImageTk
from itertools import count, cycle
import rookiepy as bc

# Setup logging
logging.basicConfig(
    filename="recent.log",
    encoding='utf-8',
    format='%(asctime)s %(levelname)-8s %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S'
    )

# Load json config file
with open("settings.json") as jsonfile:
    config = json.load(jsonfile)
departments_dict = config.get('departments', {})
pt_departments_dict = config.get('ptdepartments', {})
groups_dict= config.get('groups', {})
police_groups= config.get('police_groups', [])
supervisors = config.get('supervisors',{})

# Function to clean up strings
def replace_special_characters(text):
    text = re.sub(r'[\n\r\t]', ' ', text)
    text = re.sub(r' +', ' ', text)
    return text

# Function to scrape tickets
def get_ticket_details(ticketNumber):

    # Load cookies from Edge
    try:
        cookies = bc.edge(domains=["spiceworks.com"])
        cookie = bc.to_cookiejar(cookies)
    except:
        messageBox.showerror("Error",f"Could not load cookies from Edge. Clear cookies and login to spiceworks on Edge or Chrome.")
        exit()
    
    url = f'https://provocity.on.spiceworks.com/api/tickets/{ticketNumber}'

    # Load cookies to authenticate to spiceworks
    response = requests.get(url, cookies=cookie, verify="cert.pem")
    try:
        data = json.loads(response.text)
    except ValueError:
        #if no JSON was returned, try again with Chrome cookies
        try:
            # load chrome cookies
            try:
                cookies = bc.chrome(domains=["spiceworks.com"])
                cookie = bc.to_cookiejar(cookies)
            except:
                messageBox.showerror("Error",f"Could not load cookies from Chrome. Clear cookies and login to spiceworks on Edge or Chrome.")
                exit()
            # make request and try to parse JSON
            response = requests.get(url, cookies=cookie, verify="cert.pem")
            data = json.loads(response.text)
        except ValueError as e:
            logging.error(f"Could not parse API response. Likely did not return correct JSON format. More info: {e}")
            messageBox.showerror("API Error", "Spiceworks API did not provide correct JSON data")
            exit()

    # Get ticket body from JSON response
    rawtext=data["ticket"]["description"]
    text = replace_special_characters(rawtext)

    # Verify the ticket was a New Hire ticket and set variables
    if data["ticket"]["summary"] != "Employee Hire" or "Full-Time Employee Hire":
        employee_pattern = r"Employee: (.+?) - (\d+)"
        position_pattern = r"Position: P-\d*(?:\w*) (.*?)(?=\s\(|\s-\s| Requisition)"
        supervisor_pattern = r"Supervisor: (\w+ \w+)"
        
        employee_pattern2 = r"Name: (.+?) -"
        employeeID_pattern2 = r"EE ID: (.+?) -"
        position_pattern2 = r"Title: (.+?) -"
        supervisor_pattern2 = r"Supervisor: (.+? .+? )"

        employee_match = re.search(employee_pattern, text)
        position_match = re.search(position_pattern, text)
        supervisor_match = re.search(supervisor_pattern, text)
        employeeID_match = None

        if employee_match is None:
            employee_match = re.search(employee_pattern2, text)
            employeeID_match = re.search(employeeID_pattern2, text)
        if position_match is None:
            position_match = re.search(position_pattern2, text)
        if supervisor_match is None:
            supervisor_match = re.search(supervisor_pattern2, text)

        if employee_match and position_match and supervisor_match and employeeID_match is None:
            employeeName = employee_match.group(1).strip()
            employeeId = employee_match.group(2).strip()
            position = position_match.group(1).strip()
            supervisor = supervisor_match.group(1).strip()

            result = employeeName, employeeId, position, supervisor
            return result
        elif employee_match and employeeID_match and position_match and supervisor_match:
            employeeName = employee_match.group(1).strip()
            employeeId = employeeID_match.group(1).strip()
            position = position_match.group(1).strip()
            supervisor = supervisor_match.group(1).strip()

            result = employeeName, employeeId, position, supervisor
            return result
        else:
            messageBox.showerror("Error",f"No Ticket found for #{ticketNumber}")
            logging.error(f"No tickets found for {ticketNumber}")
            return None

# Function to start thread (to keep program responsive)
def submit():
    submit_button.config(state=tk.DISABLED)
    thread = threading.Thread(target=update_ad_record)
    thread.start()

    window.after(100, check_thread, thread)

# Function to clear out entries when program finishes
def check_thread(thread):
    if thread.is_alive():
        window.after(100, check_thread, thread)
    else:
        submit_button.config(state=tk.NORMAL)
        entry_ticket_number.delete(0, tk.END)
        needs_email_var.set(0)
        part_time_var.set(0)

# Function to get departent of given user
def get_department(user):
    q = adquery.ADQuery()

    # Get department DN and username from common name
    try:
        q.execute_query(attributes=["department", "distinguishedName","sAMAccountName"],
                        where_clause=f"cn='{user}'")
    except:
        messageBox.showerror("Error",f"Could not find a valid department for {user}")
        logging.error(f"Could not find a valid department for {user}")
        return None
    else:
        #If no users are returned, promt for correct spelling of the manager's name
        if q.get_row_count() == 0:
            while True:
                actualName = createManModal(user)
                if actualName is None:
                    return None
                try:
                    # Try query again with the correct name
                    q.execute_query(attributes=["department", "distinguishedName", "sAMAccountName"],
                                    where_clause=f"cn='{actualName}'")
                except Exception as e:
                    messageBox.showerror(f"Error: {str(e)}")
                    logging.error(str(e))
                    return None
                if q.get_row_count() == 1:
                    break 
        # If 1 user is returned, assign fields to variables
        if q.get_row_count() == 1:
            departmentName=q.get_single_result()["department"]
            supervisor_dn=q.get_single_result()['distinguishedName']
        # If more than 1 user is found, promt for correct one
        if q.get_row_count() > 1:
            all = q.get_all_results()
            all_usernames = [item['sAMAccountName'] for item in all]
            correctName = createModal(user, all_usernames)
            while correctName.get() not in all_usernames:
                time.sleep(2)
                userName = correctName.get()
            for item in all:
                if item['sAMAccountName'] == userName:
                    supervisor_dn = item['distinguishedName']
                    departmentName = item['department']
            pop.destroy()
        return departmentName, supervisor_dn

# Function to find a user from an employee ID
def get_user_from_id(employeeId):
        q = adquery.ADQuery()
        # Query for DN and username from employee ID
        try:
            q.execute_query(attributes=["distinguishedName","sAMAccountName"],
                            where_clause=f"employeeID='{employeeId}'")
        except:
            messageBox.showerror("Error",f"Could not find a user for employeeID: {employeeId}")
            return None
        # Handle exception of no account found or multiple accounts found
        else:
            if q.get_row_count() == 0:
                messageBox.showerror("Error",f"Could not find a user for employeeID: {employeeId}")
                return None
            if q.get_row_count() > 1:
                messageBox.showerror("Error",f"Multiple accounts with employee ID: {employeeId}")
                return None
            result = q.get_single_result()
            user_dn = result["distinguishedName"]
            email = result["sAMAccountName"]
            return user_dn, email

# function to set email proxies correctly in AD
def set_email(email, userObject):
    # assign proxies to variable
    email_proxies=[f"SMTP:{email}@provo.org",f"smtp:{email}@provo.utah.gov"]
    # set proxyAddresses and mail field in AD attributes
    email_dict={"proxyAddresses":email_proxies,"mail":f"{email}@provo.utah.gov"}
    # add user to ciscoduosync
    emailGroup=adgroup.ADGroup.from_cn("ciscoduosync")
    try:
        userObject.update_attributes(email_dict)
        userObject.add_to_group(emailGroup)
    except:
        messageBox.showerror("Error","Error updating email proxies and adding to email group\nCheck AD to verify")
        logging.error("Error updating email proxies and adding to email group\nCheck AD")

# Function to update AD Object with details provided in ticket
def update_ad_record():
    pythoncom.CoInitialize()

    # get user input from fields
    ticketNumber = entry_ticket_number.get()
    email_status = needs_email_var.get()
    part_time = part_time_var.get()
    try:
        #check if a ticket number was entered
        if not ticketNumber:
            messageBox.showerror("Error","Please provide a ticket number")
            return

        # Get user information from ticket details
        logging.info("getting ticket details")
        userInfo = get_ticket_details(ticketNumber)
        if userInfo is None: return
        employeeName, employeeId, position, supervisor = userInfo

        if supervisor in supervisors:
            supervisor = supervisors[supervisor]

        # Get department and manager dn from supervisor
        logging.info("getting department")
        department = get_department(supervisor) 
        if department is None: return
        departmentName, manager_dn = department

        # Find the user in Active Directory
        logging.info(f"getting user from {employeeId}")
        user = get_user_from_id(employeeId)
        if user is None: return
        user_dn, email = user
            
        # Update user attributes
        userObj = aduser.ADUser.from_dn(user_dn)
        logonCount = userObj.get_attribute("logonCount")
        attributes={"title":position,"department":departmentName,"manager":manager_dn}

        # Add police groups
        if position == "Police Officer I":
            for groupCN in police_groups:
                groupObj = adgroup.ADGroup.from_cn(groupCN)
                userObj.add_to_group(groupObj)

        #set email attributes
        if email_status == 1:
            set_email(email, userObj)

        #move to PT OU if part-time
        if part_time == 1:
            departmentOU = pt_departments_dict[departmentName]
        else:
            departmentOU = departments_dict[departmentName]
            deptGroupObj=adgroup.ADGroup.from_cn(groups_dict[departmentName])
            try:
                userObj.add_to_group(deptGroupObj)
            except:
                messageBox.showerror("Error","Error adding user to department group")

        userObj.update_attributes(attributes)
        try: userObj.move(adcontainer.ADContainer.from_dn(departmentOU))
        except: pass
        
        # Enable user account if disabled
        userObj.enable()
        try:
            if pyadutils.convert_datetime(userObj.get_attribute("accountExpires")[0]):
                messageBox.showerror("Error", "User has an expiration\nThis needs to be manually removed")
        except: pass
        
        copytext = f"The account for {employeeName} has been created with the username '{email}', please contact the IS Helpdesk at ext.6560 for the account password."
    

        #check if user has existing account
        if len(logonCount) == 0 or logonCount[0] == 0:
            password = config.get('default_pass', '')
            userObj.set_password(password)
            userObj.force_pwd_change_on_login()
            messageBox.showinfo("Result",f"User {employeeName} has been updated\nDepartment: {departmentName}\nPosition: {position}\nManager: {supervisor}\n\nRemember to Create P-Drive folder if needed")
            clipboard.copy(copytext)
        else:
            # Prevent password being reset on active accounts
            messageBox.showinfo("Result",f"User may have an active account, password has not been reset\nUser {employeeName} has been updated\nDepartment: {departmentName}\nPosition: {position}\nManager: {supervisor}\n\nRemember to Create P-Drive folder if needed")
            clipboard.copy(copytext)

        logging.info(f"{getpass.getuser()} updated User: {employeeName}. Department: {departmentName} Position: {position} Manager: {supervisor}")
    
    except Exception as e:
        logging.error(e)
        messageBox.showerror("error",f"Error: {e}")

    pythoncom.CoUninitialize()

# Function to load random gif for GUI
def get_random_gif(folder_path):
    try:
        file_list = os.listdir(folder_path)
        gif_files = [file for file in file_list if file.endswith('.gif')]
        random_gif = random.choice(gif_files)
        return random_gif
    except:
        return None
class ImageLabel(tk.Label):
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im)
        frames = []
        try:
            for i in count(1):
                frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            pass
        self.frames = cycle(frames)
        try:
            self.delay = im.info['duration']
        except:
            self.delay = 100

        if len(frames) == 1:
            self.config(image=next(self.frames))
        else:
            self.next_frame()
    def unload(self):
        self.config(image=None)
        self.frames = None
    def next_frame(self):
        if self.frames:
            self.config(image=next(self.frames))
            self.after(self.delay, self.next_frame)

# Modal for choosing correct supervisor
def createModal(name = str, options = list):
    global pop
    pop = tk.Toplevel(window)
    pop.title("Action Needed")
    pop.option_add("*Font", comic_sans_font)
    pop.wm_iconphoto(False, icon)

    label = tk.Label(pop, text=f"Multiple supervisors found for {name}.\nChoose the correct one")
    label.pack()
    v = tk.StringVar()
    for text in options:
        tk.Radiobutton(pop, text = text, variable = v,
                    value = text, indicator = 0,
                    background = "light blue").pack(fill = tk.X, ipady = 5)
    return v

# Modal for multiple users found with a given name
def createManModal(user = str):
    newWin = tk.Tk()
    newWin.withdraw()
    input = simpledialog.askstring("Manual Entry", f"Enter correct name as it appears in AD for {user}", parent=newWin)
    newWin.destroy()
    return input

#create tkinter window
window = tk.Tk()
window.title("Update Active Directory Record")
comic_sans_font = tkfont.Font(family="Comic Sans MS", size=12)
window.option_add("*Font", comic_sans_font)
icon = tk.PhotoImage(file="logo.png")
window.wm_iconphoto(False, icon)

#structure tkinter window
lbl_ticket_number = tk.Label(window, text="Ticket Number:")
entry_ticket_number = tk.Entry(window)

needs_email_var = tk.IntVar()
email_checkbox = tk.Checkbutton(window, text="Needs email?", variable=needs_email_var)

part_time_var = tk.IntVar()
part_time_checkbox = tk.Checkbutton(window, text="Part time employee?", variable=part_time_var)

submit_button = tk.Button(window, text="Submit", command=submit)

lbl = ImageLabel(window)
lbl.pack()
randomgif = get_random_gif("gifs/")
if randomgif is not None:
    lbl.load(f'gifs/{randomgif}')

lbl_ticket_number.pack()
entry_ticket_number.pack()

email_checkbox.pack()

part_time_checkbox.pack()

submit_button.pack()

#start program
window.mainloop()
