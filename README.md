# ***TermServ+***

***Version:** 1.2.0*

***Author:** Joshua Dwight ([joshdwight101](https://github.com/joshdwight101))*

![screenshot](screenshots/Screenshot%202026-04-07%20103118.png)

*TermServ+ is an advanced, lightweight PowerShell application featuring a native C\# Windows Forms GUI designed to supercharge the standard Windows Remote Desktop Connection (MSTSC) experience. Built specifically for systems administrators, IT professionals, and enterprise environments, TermServ+ eliminates the friction of managing multiple server connections and repetitive authentication prompts.*

*The application compiles entirely in-memory at runtime, requiring no complex installations, external dependencies, or executable binaries that might flag enterprise security software.*

## ***🚀 The TermServ+ Advantage (Exclusive Features)***

*While the standard MSTSC client is great for one-off connections, it lacks bulk management capabilities. TermServ+ introduces several powerful features designed to save time and streamline workflows:*

* ***Global Credential Deployment:** Mass-deploy your login credentials to the Windows Credential Manager (cmdkey) for every server in your list simultaneously. Connect to any server on your list without seeing a login prompt. When your password eventually expires, simply type the new one here and update your credentials across your entire infrastructure in a single click.*  
* ***One-Click Credential Clearing:** Instantly wipe all saved credentials for your server list from the Windows Credential Manager when security policies require a complete session termination or credential flush.*  
* ***Portable Server List Management:** Servers are automatically saved to a lightweight TermServPlus.ini file in the same directory as the application.*  
* ***Intelligent Auto-Save:** The application continually saves your configurations in the background as you type, change settings, or switch tabs. No manual saving required.*  
* ***Auto-Populating Context:** Automatically detects and populates your current Windows Domain and Username to speed up the credential saving process. Includes a quick-clear (X) button for easy overriding.*  
* ***Password Visibility Toggle:** A built-in mask/unmask checkbox allows you to verify your complex passwords before deploying them to the Credential Manager.*  
* ***Rapid Tab Navigation:** Use keyboard shortcuts (Ctrl+1, Ctrl+2, Ctrl+3, Ctrl+4) to rapidly switch between configuration tabs without touching your mouse.*

## ***🏢 Enterprise Power Use & Intune Deployment***

*The core philosophy behind TermServ+ is to make enterprise RDP access as seamless as possible for end-users and administrators.*

*Because TermServ+ relies on an external TermServPlus.ini file, administrators can:*

1. ***Pre-configure the INI:** Create a master TermServPlus.ini containing all required infrastructure servers and per-server RD Gateway configurations.*  
2. ***Sign the Script (Security Best Practice):** Digitally sign the TermServPlus.ps1 script using your organization's internal PKI or a trusted Code Signing certificate (Set-AuthenticodeSignature). This ensures the script runs natively under strict enterprise AllSigned or RemoteSigned execution policies without warnings.*  
3. ***Package for Intune/MECM:** Package the signed TermServPlus.ps1 and the pre-configured TermServPlus.ini together as an .intunewin package.*  
4. ***Deploy:** Deploy the package to end-user machines (e.g., C:\\ProgramData\\TermServ+) and push a desktop shortcut targeting the script.*  
5. ***Painless Password Rotations:** When an end-user opens the app, they simply type their password once on the General tab and click "Save Creds". They can then instantly connect to any server in their assigned infrastructure without being prompted for credentials. When mandatory password expirations occur, the user only has to update their password in TermServ+ once to regain seamless access to everything.*

## ***🖥️ Standard MSTSC Feature Parity***

*Under the hood, TermServ+ dynamically generates customized .rdp configurations and pipes them directly into the native Windows mstsc.exe engine. This ensures 100% compatibility with standard RDP protocols and features:*

* ***Display Settings:***  
  * *Configurable resolutions (Full Screen up to custom window sizes).*  
  * *Multi-monitor support.*  
  * *Color depth configuration (15-bit to 32-bit Highest Quality).*  
* ***Local Resources:***  
  * ***Audio:** Configure remote audio playback and recording redirection.*  
  * ***Keyboard Hooks:** Apply Windows key combinations (e.g., Alt+Tab) locally, remotely, or only in full screen.*  
  * ***Devices:** Redirect local printers and clipboard.*  
* ***Advanced Connections:***  
  * ***Console Access:** Connect directly to the administrative/console session of a server.*  
  * ***Per-Server RD Gateway Settings:** Robust, fully-featured RD Gateway configurations saved independently for each server in your list. Includes support for auto-detection, specific server assignment, logon methods (NTLM/Smartcard), local bypass rules, and credential reuse.*

## ***🛠️ Usage & Installation***

1. *Clone or download this repository.*  
2. ***Execution Policy:** Ensure your PowerShell environment allows script execution. In enterprise environments, it is highly recommended to **digitally sign** the TermServPlus.ps1 script with a trusted certificate rather than lowering your machine's global execution policy (e.g., bypassing with Set-ExecutionPolicy RemoteSigned \-Scope CurrentUser).*  
3. *Right-click TermServPlus.ps1 and select **Run with PowerShell**, or execute it from your preferred terminal.*  
4. *The application will automatically generate a template TermServPlus.ini file upon its first run if one does not exist.*

*Created by Joshua Dwight. Feel free to fork, contribute, and enhance your infrastructure management.*

![screenshot](screenshots/Screenshot%202026-04-07%20103118.png)
![screenshot](screenshots/Screenshot%202026-04-07%20103126.png)
![screenshot](screenshots/Screenshot%202026-04-07%20103133.png)
![screenshot](screenshots/Screenshot%202026-04-07%20103138.png)
