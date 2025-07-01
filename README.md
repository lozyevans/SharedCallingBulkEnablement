# Teams Shared Calling Enablement Script

This document provides a step-by-step guide to setting up and running the `sharedCalling_enablement_v1.1.ps1` script. It includes details on pre-requisites, configuration, permissions, and usage of the script, including its menu system and logging.

The script can be used to enable and configure Shared Calling functionality in Microsoft Teams, allowing multiple users to share the same phone number and receive calls to that number.

Note: The script should be run from a location that has line of sight to Microsoft 365 cloud services.

---

## Pre-requisites

### 1. PowerShell Modules
Ensure the following PowerShell modules are installed:
- **MicrosoftTeams**: Used for managing Teams and related resources.
- **Microsoft.Graph**: For managing Microsoft Graph API operations.

#### Checking Installed Modules
Run the following commands to verify the modules are installed:

```powershell
# Check if MicrosoftTeams module is installed
Get-Module -Name MicrosoftTeams -ListAvailable

# Check if AzureAD module is installed
Get-Module -Name AzureAD -ListAvailable

# Check if MS Graph module is installed
# If you have more than one listed, you will need to remediate, follow this guide https://bonguides.com/how-to-fix-get-mguser-one-or-more-errors-occurred/
Get-Module Microsoft.Graph -ListAvailable
```

---

### 2. Permissions
Ensure the following permissions are granted to the account running the script:
- **Global Administrator** or **Teams Administrator** role in Azure AD.
- **Teams Phone Administrator** role for managing phone-related tasks.

---

### 3. Configuration
Before running the script, configure the following settings in the `config.xml` file. This file contains all the necessary parameters for the script to function correctly.
An example is provided below:

#### Example `config.xml`
```xml
```xml
<?xml version="1.0"?>
<settings>
<EntraID></EntraID>
  <ServicePlanName>MCOEV,MCOPSTN_PAYG_1</ServicePlanName> <!-- Service Plan Name, Direct Routing requires E5 (MCOEV) or, Microsoft Calling Plans require PAYG (MCOPSTN_PAYG_1) and Comms credits (MCOPSTNC). If multiple licenses, comma separate them -->
  <SendChat>True</SendChat>
  <ChatImage>https://hybridmeetingrooms.z33.web.core.windows.net/TeamsSharedCalling3.png</ChatImage>
  <MigrationSuccessMessage>You have been successfully enabled for Teams Shared Calling.</MigrationSuccessMessage>
  <MigrationFailMessage>Teams Shared Calling enablement failed. Please wait for further updates.</MigrationFailMessage>
   <TenantDialPlan>
    <GroupName></GroupName>
    <PolicyName>MNO-DP-GB</PolicyName>
  </TenantDialPlan>
  <OnlineVoiceRoutingPolicy>
    <GroupName></GroupName>
    <PolicyName>TVRP_GB_All_Unrestricted</PolicyName>
  </OnlineVoiceRoutingPolicy>
  <CallingLineIdentity>
    <GroupName></GroupName>
    <PolicyName>CLI_UK_COR_Main</PolicyName>
  </CallingLineIdentity>
  <TeamsCallingPolicyUsers>
    <GroupName></GroupName>
    <PolicyName>CP_Call_Agents</PolicyName>
  </TeamsCallingPolicyUsers>
  <TeamsVoicemailPolicy>
    <GroupName></GroupName>
    <PolicyName>TranscriptionProfanityMaskingEnabled</PolicyName>
  </TeamsVoicemailPolicy>
  <TeamsSharedCallingPolicy>
    <GroupName>ASG-TMS-U-SCP_UK_COR_44300012345</GroupName>
    <PolicyName>SCP_UK_COS_44300012345</PolicyName>
  </TeamsSharedCallingPolicy>
  <TeamsEmergencyCallingPolicy>
    <GroupName></GroupName>
    <PolicyName>ECP_UK_COR</PolicyName>
  </TeamsEmergencyCallingPolicy>
  <TeamsEmergencyCallRoutingPolicy>
    <GroupName></GroupName>
    <PolicyName>ECRP_UK_COR_999</PolicyName>
  </TeamsEmergencyCallRoutingPolicy>
  <TeamsVoiceApplicationPolicy>
    <GroupName></GroupName>
    <PolicyName>TVA_AACQ_01</PolicyName>
  </TeamsVoiceApplicationPolicy>
```
</settings>
```

#### Explanation of `config.xml` Elements
- **EntraID**: The Azure Active Directory (AAD) tenant ID or domain name. This can be left blank.
- **ServicePlanName**: The name of the service plan to check for licensing (e.g., `MCOEV` for E5 Teams Phone, `MCOPSTN_PAYG_1` for pay-as-you-go calling plans). Multiple licenses can be comma-separated.
- **SendChat**: To send the user a Teams Chat message indicating their shared calling enablement status, set this to True. *See notes.
- **ChatImage**: A URL to the image that will be included in the Teams Chat message.
- **MigrationSuccessMessage**: The message sent to users when shared calling enablement succeeds.
- **MigrationFailMessage**: The message sent to users when shared calling enablement fails.
- **TenantDialPlan**: Contains the `GroupName` and `PolicyName` for the Tenant Dial Plan. Use `GroupName` for group-based policy assignment or `PolicyName` for direct assignment.
- **OnlineVoiceRoutingPolicy**: Contains the `GroupName` and `PolicyName` for the Online Voice Routing Policy.
- **CallingLineIdentity**: Contains the `GroupName` and `PolicyName` for the Calling Line Identity Policy.
- **TeamsCallingPolicyUsers**: Contains the `GroupName` and `PolicyName` for the Teams Calling Policy for users.
- **TeamsVoicemailPolicy**: Contains the `GroupName` and `PolicyName` for the Teams Voicemail Policy.
- **TeamsSharedCallingPolicy**: Contains the `GroupName` and `PolicyName` for the Teams Shared Calling Policy. This is essential for shared calling functionality.
- **TeamsEmergencyCallingPolicy**: Contains the `GroupName` and `PolicyName` for the Teams Emergency Calling Policy.
- **TeamsEmergencyCallRoutingPolicy**: Contains the `GroupName` and `PolicyName` for the Teams Emergency Call Routing Policy.
- **TeamsVoiceApplicationPolicy**: Contains the `GroupName` and `PolicyName` for the Teams Voice Application Policy.

> **Note**: For each policy, either `GroupName` or `PolicyName` should be used, but not both. If both are specified, the script will prioritize `GroupName`. All policy names and group names must match exactly as they appear in Teams/Entra ID.

> **Note**: 
- The chat functionality has a requirement on the Microsoft Graph PowerShell modules. Use the script named 'Fix-Graph.ps1' located in the root folder to repair your terminal. This script will uninstall any Microsoft Graph modules currently installed and then reinstall them with the latest versions. This can take between 2-3 hours to complete. 
- For each policy, either `GroupName` or `PolicyName` should be used, but not both. If both are specified, the script will prioritise `GroupName`. Any policyName or GroupName added must match exactly against the policy name or Group name in Teams/Entra. Only add names into the policy assignments when **not** using a Global policy assignment.

### The batch file
The batch file, which should be placed in the data folder, is a CSV file that contains user-specific information required for the shared calling enablement process. 
This file must have the following required headers:

UPN: The User Principal Name (UPN) of the user. This is the unique identifier for the user in Entra ID and must be in the format of an email address (e.g., user@contoso.com).
DDI: Add the users Extenion number here that is to be assigned in Entra, in the format of xNNNN for example x11234 (the extension number can be as long as you specifiy but must start with an x)

Placement:
Save the CSV file in the **data** folder within the root directory of the script.
Ensure the file is named appropriately (e.g., shared_calling_batch1.csv) so it can be easily identified when running the script.

---

## Script File Structure
There are a number of folders which are used to either store logs, export data, contain user batch files or as a temporary store.
The folder structure is as follows:
root
├── arc
├── data
├── exp
├── logs
└── tmp


1. root
This is the base directory where all other folders reside. It contains the sharedCalling_enablement_v1.1.ps1 script, config.xml and other essential files like the readme.md file.
The script references subfolders within this directory for organizing input/output data or temporary files.
2. arc
Purpose: Archive folder.
This folder will store old or completed enablement data, such as logs or exported files from previous runs of the script. It helps maintain a history of operations for auditing purposes.
3. data
Purpose: Input data folder.
This folder is used to contain the input files required by the script, such as user batches. The script will read from this folder to determine which users need to be enabled for shared calling.
4. exp
Purpose: Export folder.
This folder will store exported data generated during the enablement process. For example, the script will export status reports for user enablement, telephone number assignment.
5. logs
Purpose: Log storage.
This folder is used to store log files generated by the script. These logs can help track the progress of the enablement, identify errors, and debug issues if something goes wrong.
6. tmp
Purpose: Temporary files.
This folder is used for temporary files created during the script's execution. These files might include intermediate data, temporary exports, or other transient information that is cleaned up after the script completes.

---

## Script Structure
The `sharedCalling_enablement_v1.1.ps1` script is modular and organized into multiple functions, each designed to perform a specific task. This structure ensures that the script is easy to maintain, extend, and debug. Below is an overview of the key components:

### 1. Initialization
The script begins by loading required modules, validating permissions, and reading the `config.xml` file. It ensures that all pre-requisites are met before proceeding.

### 2. Menu System
The script includes an interactive menu system that allows users to select specific tasks to perform. Each menu option corresponds to a function within the script. The menu system is designed to guide users through the shared calling enablement process step-by-step.

### 3. Core Functions
The script is divided into several core functions, each responsible for a specific operation:
- **EnableSharedCalling**: Enables shared calling for users specified in the batch file by assigning the Shared Calling Routing Policy.
- **ConfigureCallingPolicies**: Applies necessary calling policies for shared calling.
- **TeamsUserStatus**: Exports a detailed report of Teams user status, including assigned policies and attributes.
- **ApplyPolicies**: Assigns Teams policies such as Calling Policies, Dial Plans, and Voice Routing Policies to users.

### 4. Logging
Each function generates detailed logs that are stored in the `logs` folder. These logs include timestamps, operation details, and error messages (if any), making it easier to troubleshoot issues.

### 5. Configuration Handling
The script reads the `config.xml` file to retrieve settings such as policy names, group names, and service configurations. It validates the configuration before proceeding with any operation.

### 6. Error Handling
The script includes robust error handling to ensure that issues are logged and do not disrupt the overall execution. If an error occurs, the script provides meaningful error messages and suggests corrective actions and is appended to the relevant debug log for that function.

### 7. Cleanup
Temporary files created during the script's execution are stored in the `tmp` folder and are cleaned up automatically after the script completes.

This modular structure ensures that the script is flexible, user-friendly, and capable of handling shared calling enablement scenarios efficiently.


## Usage

### Running the Script
1. Open PowerShell 5.1.
2. Navigate to the directory containing the script.
3. Execute the script using the following command:
    ```powershell
    .\sharedCalling_enablement_v1.1.ps1
    ```
    The Menu will be displayed.

### Menu System
The script provides an interactive menu system with the following options:

**File Operations:**
- **Option A**: Archive (Current) - Archive current logs and exports
- **Option B**: Batch File - Import a batch file for processing
- **Option C**: Show Config File - Display current configuration
- **Option D**: Clear debugs - Remove debug log files

**Pre-Checks:**
- **Option 0**: Check Microsoft Graph Status - Verify Graph API connectivity
- **Option 1**: Licensing status - Check user licensing for Teams Phone
- **Option 2**: Check Policy Assignment - Verify policy assignments in the tenant
- **Option 3**: Teams user status - Review Teams configuration for users

**Enterprise Voice:**
- **Option 10**: Enable Users for Enterprise Voice - Configure enterprise voice features
- **Option 11**: Apply User Extension Numbers - Assign extension numbers to users

**Assign Teams Policies:**
- **Option 12**: Apply Tenant Dial Plan Policy - Assign dial plans to users
- **Option 13**: Apply Online Voice Routing Policy - Configure voice routing
- **Option 14**: Apply Teams Calling Policy - Set calling features and permissions
- **Option 15**: Apply Caller ID Policy - Configure caller ID settings
- **Option 16**: Apply Teams Shared Calling Policy - Enable shared calling functionality
- **Option 17**: Apply Teams Emergency Calling Policy - Set emergency calling parameters
- **Option 18**: Apply Teams Emergency Call Routing Policy - Configure emergency call routing
- **Option 19**: Apply Teams Voice Application Policy - Set voice application parameters

**Inform Users:**
- **Option 20**: Send Bulk Chat - Message users with shared calling information

**Misc:**
- **Option H**: Help - Display help information
- **Option Q**: Quit - Exit the script

### Recommended Workflow
For the most effective use of this script, follow this logical workflow:

1. **Initial Setup**:
   - **Option B**: Import your batch file first to define which users will be enabled
   - **Option C**: Review your configuration settings to ensure they match your requirements
   - **Option 0**: Verify Microsoft Graph connectivity before proceeding

2. **Pre-Checks**:
   - **Options 1-3**: Run through all pre-checks to verify licensing and existing configurations
   - Pay special attention to Option 1 to ensure users have appropriate Teams Phone licensing

3. **Enable Core Voice Features**:
   - **Option 10**: Enable Enterprise Voice for the users
   - **Option 11**: Apply extension numbers to users

4. **Apply Required Policies**:
   - **Options 12-15**: Apply dial plan, voice routing, and calling policies
   - **Option 16**: Apply the Teams Shared Calling Policy (critical for shared calling functionality)
   - **Options 17-19**: Apply emergency calling policies and voice application settings

5. **Complete the Process**:
   - **Option 20**: Notify users about their newly enabled shared calling capabilities
   - **Option A**: Archive logs and exports for record-keeping
   - **Option D**: Clear debug logs if needed

This structured approach ensures that all necessary components are correctly configured for shared calling functionality.

---

## Additional Information

### Additional Options
The script provides the following additional options in the menu:

1. **Check Group-Based Policy Assignment**:  
    Verifies if group-based policy assignments exist in Azure AD. Run this as a pre-check before applying policies when using Group-Based Policy Assignment.

---

### Usage Instructions
Follow these steps to use the script effectively:

1. **Ensure Prerequisites**:  
    Make sure you have the necessary permissions and prerequisites set up before running the script. This includes installing the required PowerShell modules:  
    - MicrosoftTeams  
    - Microsoft.Graph  

2. **Configure `config.xml`**:  
    Apply the relevant settings within the `config.xml` file located in the root directory. Ensure to:  
    - For each policy to be explicitly assigned (i.e., not a global policy), provide the Teams Policy name or the Group Name if using Group-Based Policy Assignment. These names must match exactly.

3. **Run the Script**:  
    Execute the script, which will load the main menu.

4. **Import the Batch File**:  
    Start by selecting **Option B** to import the batch file. When prompted, add a Group Identifier or Batch Number for the enablement.

5. **Perform Pre-Checks**:  
    Run through the environmental pre-checks using the appropriate menu options.

6. **Perform Desired Actions**:  
    Use the options in the menu to perform the desired actions, such as:  
    - Enable shared calling  
    - Apply policies  
    - Check Teams status  

7. **Archive Logs and Exports**:  
    Use **Option A** to archive logs and exports after completing your tasks. This step is optional but recommended for keeping track of changes and any errors encountered.

8. **Clear Debug Logs**:  
    Use **Option C** to clear debug logs if needed.

9. **Exit the Script**:  
    To exit the script, select **Option Q**.

### Considerations

- Users must be licensed with Phone System before they can be enabled for shared calling.
- The Teams Chat capability can be temperamental due to the way the Microsoft Graph modules interact and can throw an exception stating the chatbody is missing. In this case, the Fix-Graph.ps1 script will need to be run, which can take up to 3hrs to perform, or update the config file to disable it.
- Consider chunking the batch files into batches of 1000 users per batch.
- For different shared calling groups, have separate config.xml files. To change the config file, simply rename it to config.xml to be loaded into the script.

## Logging
The script generates detailed logs for each operation. Logs are stored in the directory specified in the `config.xml` file. Ensure the directory exists and is writable before running the script.

## Known Issues

There is currently one known issue:
1. If you encounter an issue with the chat function and it throws an exception as per below

    New-MgChat : Provided payload is invalid. Errors -
    Key - 'ChatType' , Error - 'The ChatType field is required.'
    Status: 400 (BadRequest)
    ErrorCode: BadRequest
    Date: 2025-05-07T16:36:44
    Headers:
    Transfer-Encoding             : chunked
    Vary                          : Accept-Encoding
    Strict-Transport-Security     : max-age=31536000
    request-id                    : efaed9f4-ea22-4872-ba40-54d63ea5981f
    client-request-id             : 62ad8087-66e0-45c4-8d9d-72196c192cca
    x-ms-ags-diagnostic           : {"ServerInfo":{"DataCenter":"UK
    South","Slice":"E","Ring":"5","ScaleUnit":"003","RoleInstance":"LO2PEPF00003339"}}
    Date                          : Wed, 07 May 2025 16:36:44 GMT
    At line:1 char:1

    This is likely due to the MgGraph module version.

    To resolve this issue run the Fix-MgGraph.ps1 script however update the two install lines at the end of the script to use a known working module version:
    
    Install-Module Microsoft.Graph -RequiredVersion 2.25.0 -Force
    Install-Module Microsoft.Graph.Beta -RequiredVersion 2.25.0 -Force

