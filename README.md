# Cerberus AEP Result Script

## Google API Setup
A Google Cloud Project is necessary for this project. You can create a project through the [Google Cloud Dashboard](https://console.cloud.google.com/home/dashboard).

### 1. Enable APIs
Once you have your project created, it's necessary to enable the **Google Drive API** and **Google Sheet API**.
+ Using the search bar at the top search for and enable the Google Drive API and Google Sheet API modules.

![alt text](https://github.com/1aurelius/AEP-Result-Script/blob/main/images/google_sheets_api.png)
![alt text](https://github.com/1aurelius/AEP-Result-Script/blob/main/images/google_drive_api.png)

### 2. Create a Service Account
+ Navigate to the **APIs & Services** page.
+ From there, select the **Create Credentials** button and click **Service account**.

![alt text](https://github.com/1aurelius/AEP-Result-Script/blob/main/images/create_service_account.png)
+ Then, fill out the form and click **DONE**.

### 3. Create a Service Account Key
+ Once your service account is created, click on the **Manage service accounts** button next to the Service Accounts section.
+ From there, click on the **⁝** button under "Actions" and select **Manage Keys**.

![alt text](https://github.com/1aurelius/AEP-Result-Script/blob/main/images/manage_keys.png)
+ After that, select the **"ADD KEY"** button and create a new key.

![alt text](https://github.com/1aurelius/AEP-Result-Script/blob/main/images/create_key.png)
+ Make sure that the key type is JSON and then click **CREATE**.
+ Your key will then be downloaded. Make sure to remember where you store this key. I recommend you store it in the same folder as the Python file.

## Google Sheet Setup
Once you've set up all of the API modules necessary for the project you can move on to setting up a Google Sheet. 

### 1. Create a Mirror
Since we don't want the API to directly access the main AEP Database, we need to create a mirror that the API can use instead.
+ Create a blank Google Sheet. Name it whatever you want.
+ In cell A1 paste `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1K5c3F7GYqk67REy2gnA6YCTd90iHhY739Dx2zBDj_vQ/edit#gid=94805241", "AEP Responses!A7:V")`. If Google Sheets asks you to authorize access, go ahead and do so.

### 2. Add the Service Account as an Editor
The Service Account needs to be added as an editor to the mirror in order to view the data.
+ Return to your Google Cloud dashboard and select the **IAM & Admin** menu. From there, navigate to the **Service Accounts** page.
+ Copy the service account email.
+ Add the service account email as an editor on the mirror.

![alt text](https://github.com/1aurelius/AEP-Result-Script/blob/main/images/share_with_service_account.png)

## Configure main.py
Some variables have to be changed in the python script in order for it to run correctly.
+ Insert the path for your service account key into the `key_path` variable.
+ Insert your rank in the `rank` variable.
+ Insert your username in the `username` variable.
