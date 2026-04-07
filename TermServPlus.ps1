# ==============================================================================
# Script: TermServPlus.ps1
# Author: Joshua Dwight
# Description: A C# GUI built inside PowerShell to manage RDP server lists 
#              and deploy Windows Credential Manager entries for MSTSC.
# ==============================================================================

# Ensure we get the correct directory whether run from ISE, VSCode, or Console
$scriptRoot = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptRoot)) { 
    $scriptRoot = (Get-Location).Path 
}

# Use a single-quoted here-string (@' ... '@) so PowerShell doesn't try to parse C# string interpolation ($"")
$csharpCode = @'
using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;

namespace TermServPlusApp
{
    public class TermServForm : Form
    {
        // Add a public property to satisfy PowerShell's Add-Type compiler and suppress the warning
        public string AppVersion { get { return "1.2.0"; } }

        private string basePath;
        private string iniPath;

        // UI Elements
        private TabControl tabMain;
        private ListBox lstServers;
        private TextBox txtNewServer;
        private Button btnAddServer;
        private Button btnRemoveServer;
        private Button btnSaveIni;
        private Button btnRefreshIni;
        private TextBox txtUsername;
        private TextBox txtPassword;
        private Button btnSaveCreds;
        private Button btnClearCreds;
        
        private ComboBox cmbResolution;
        private ComboBox cmbColors;
        private CheckBox chkMultiMon;
        
        // New Resources UI
        private ComboBox cmbAudioPlayback;
        private CheckBox chkAudioRecord;
        private ComboBox cmbKeyboard;
        private CheckBox chkPrinters;
        private CheckBox chkClipboard;
        
        // New Advanced UI
        private CheckBox chkAdmin;
        private GroupBox grpGateway;
        private Label lblSelectedGatewayServer;
        private RadioButton rdoGwAuto;
        private RadioButton rdoGwUse;
        private RadioButton rdoGwNone;
        private TextBox txtGateway;
        private ComboBox cmbGwLogon;
        private CheckBox chkGwBypass;
        private CheckBox chkGwReuseCreds;

        private Button btnConnect;
        private Button btnClose;
        private LinkLabel lnkGithub;

        // Data Storage
        private Dictionary<string, string> dictGateways = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private bool isUpdatingGateway = false;

        public TermServForm(string rootDirectory)
        {
            this.basePath = rootDirectory;
            this.iniPath = Path.Combine(basePath, "TermServPlus.ini");
            InitializeComponent();
            LoadIni();

            // Auto-save INI on program exit
            this.FormClosing += (s, e) => SaveIni();
        }

        private void InitializeComponent()
        {
            this.Text = "TermServ+ v1.2.0 by Joshua Dwight";
            this.Size = new Size(560, 640);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

            // Enable form-level keyboard shortcut listening
            this.KeyPreview = true;
            this.KeyDown += (s, e) => {
                if (e.Control)
                {
                    if (e.KeyCode == Keys.D1 && tabMain.TabPages.Count > 0) { tabMain.SelectedIndex = 0; e.Handled = true; e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.D2 && tabMain.TabPages.Count > 1) { tabMain.SelectedIndex = 1; e.Handled = true; e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.D3 && tabMain.TabPages.Count > 2) { tabMain.SelectedIndex = 2; e.Handled = true; e.SuppressKeyPress = true; }
                    else if (e.KeyCode == Keys.D4 && tabMain.TabPages.Count > 3) { tabMain.SelectedIndex = 3; e.Handled = true; e.SuppressKeyPress = true; }
                }
            };

            tabMain = new TabControl { Location = new Point(10, 10), Size = new Size(520, 520) };

            // --- TAB: GENERAL ---
            TabPage tabGeneral = new TabPage("General");
            
            Label lblServers = new Label { Text = "Server List:", Location = new Point(15, 15), AutoSize = true };
            lstServers = new ListBox { Location = new Point(15, 35), Size = new Size(220, 240) };
            lstServers.DoubleClick += BtnConnect_Click;
            lstServers.SelectedIndexChanged += LstServers_SelectedIndexChanged;

            txtNewServer = new TextBox { Location = new Point(15, 285), Size = new Size(220, 25) };
            
            btnAddServer = new Button { Text = "Add", Location = new Point(15, 315), Size = new Size(105, 30) };
            btnAddServer.Click += BtnAddServer_Click;

            btnRemoveServer = new Button { Text = "Remove", Location = new Point(130, 315), Size = new Size(105, 30) };
            btnRemoveServer.Click += BtnRemoveServer_Click;

            btnSaveIni = new Button { Text = "Save List", Location = new Point(15, 355), Size = new Size(105, 30), BackColor = Color.LightYellow };
            btnSaveIni.Click += BtnSaveIni_Click;

            btnRefreshIni = new Button { Text = "Refresh List", Location = new Point(130, 355), Size = new Size(105, 30), BackColor = Color.LightYellow };
            btnRefreshIni.Click += BtnRefreshIni_Click;

            GroupBox grpCreds = new GroupBox { Text = "Global Credentials", Location = new Point(255, 15), Size = new Size(240, 210) };
            Label lblInfo = new Label { Text = "Applies credentials to ALL servers\nvia Credential Manager.", Location = new Point(15, 25), Size = new Size(210, 35), ForeColor = Color.DimGray };
            
            Label lblUser = new Label { Text = "Username (Domain\\User):", Location = new Point(15, 65), AutoSize = true };
            txtUsername = new TextBox { Location = new Point(15, 85), Size = new Size(170, 25) };
            
            Button btnClearUser = new Button { Text = "X", Location = new Point(190, 84), Size = new Size(25, 23), BackColor = Color.LightGray };
            btnClearUser.Click += (s, e) => txtUsername.Clear();

            Label lblPass = new Label { Text = "Password:", Location = new Point(15, 115), AutoSize = true };
            txtPassword = new TextBox { Location = new Point(15, 135), Size = new Size(145, 25), PasswordChar = '*' };
            
            CheckBox chkShowPass = new CheckBox { Text = "Show", Location = new Point(165, 137), AutoSize = true };
            chkShowPass.CheckedChanged += (s, e) => {
                txtPassword.PasswordChar = chkShowPass.Checked ? '\0' : '*';
            };
            
            btnSaveCreds = new Button { Text = "Save Creds", Location = new Point(15, 165), Size = new Size(100, 30), BackColor = Color.LightGreen };
            btnSaveCreds.Click += BtnSaveCreds_Click;
            
            btnClearCreds = new Button { Text = "Clear All", Location = new Point(120, 165), Size = new Size(95, 30), BackColor = Color.LightCoral };
            btnClearCreds.Click += BtnClearCreds_Click;
            
            grpCreds.Controls.AddRange(new Control[] { lblInfo, lblUser, txtUsername, btnClearUser, lblPass, txtPassword, chkShowPass, btnSaveCreds, btnClearCreds });
            
            // Auto-populate with current domain and username
            txtUsername.Text = string.Format("{0}\\{1}", Environment.UserDomainName, Environment.UserName);
            
            tabGeneral.Controls.AddRange(new Control[] { lblServers, lstServers, txtNewServer, btnAddServer, btnRemoveServer, btnSaveIni, btnRefreshIni, grpCreds });

            // --- TAB: DISPLAY ---
            TabPage tabDisplay = new TabPage("Display");
            
            Label lblRes = new Label { Text = "Remote desktop size:", Location = new Point(20, 20), AutoSize = true };
            cmbResolution = new ComboBox { Location = new Point(20, 45), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbResolution.Items.AddRange(new string[] { "Full Screen", "1920x1080", "1600x1200", "1280x1024", "1024x768", "800x600" });
            cmbResolution.SelectedIndex = 0;
            
            chkMultiMon = new CheckBox { Text = "Use all my monitors for the remote session", Location = new Point(20, 85), Size = new Size(300, 25) };

            Label lblColors = new Label { Text = "Colors:", Location = new Point(20, 130), AutoSize = true };
            cmbColors = new ComboBox { Location = new Point(20, 155), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbColors.Items.AddRange(new string[] { "High Color (15 bit)", "High Color (16 bit)", "True Color (24 bit)", "Highest Quality (32 bit)" });
            cmbColors.SelectedIndex = 3;

            tabDisplay.Controls.AddRange(new Control[] { lblRes, cmbResolution, chkMultiMon, lblColors, cmbColors });

            // --- TAB: LOCAL RESOURCES ---
            TabPage tabResources = new TabPage("Local Resources");
            
            // Audio settings
            GroupBox grpAudio = new GroupBox { Text = "Remote audio", Location = new Point(20, 15), Size = new Size(460, 110) };
            Label lblAudio = new Label { Text = "Remote audio playback:", Location = new Point(15, 25), AutoSize = true };
            cmbAudioPlayback = new ComboBox { Location = new Point(15, 45), Size = new Size(250, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbAudioPlayback.Items.AddRange(new string[] { "Bring to this computer", "Play on remote computer", "Do not play" });
            cmbAudioPlayback.SelectedIndex = 0;
            chkAudioRecord = new CheckBox { Text = "Record from this computer", Location = new Point(15, 75), AutoSize = true };
            grpAudio.Controls.AddRange(new Control[] { lblAudio, cmbAudioPlayback, chkAudioRecord });

            // Keyboard settings
            GroupBox grpKeyboard = new GroupBox { Text = "Keyboard", Location = new Point(20, 135), Size = new Size(460, 80) };
            Label lblKeyboard = new Label { Text = "Apply Windows key combinations (e.g. ALT+TAB):", Location = new Point(15, 25), AutoSize = true };
            cmbKeyboard = new ComboBox { Location = new Point(15, 45), Size = new Size(250, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbKeyboard.Items.AddRange(new string[] { "On this computer", "On the remote computer", "Only when using the full screen" });
            cmbKeyboard.SelectedIndex = 2; // Default to full screen
            grpKeyboard.Controls.AddRange(new Control[] { lblKeyboard, cmbKeyboard });

            // Devices settings
            GroupBox grpDevices = new GroupBox { Text = "Local devices and resources", Location = new Point(20, 225), Size = new Size(460, 100) };
            Label lblDevicesInfo = new Label { Text = "Choose the devices and resources that you want to use in your remote session.", Location = new Point(15, 25), Size = new Size(400, 35) };
            chkPrinters = new CheckBox { Text = "Printers", Location = new Point(15, 65), AutoSize = true, Checked = true };
            chkClipboard = new CheckBox { Text = "Clipboard", Location = new Point(105, 65), AutoSize = true, Checked = true };
            grpDevices.Controls.AddRange(new Control[] { lblDevicesInfo, chkPrinters, chkClipboard });
            
            tabResources.Controls.AddRange(new Control[] { grpAudio, grpKeyboard, grpDevices });

            // --- TAB: ADVANCED ---
            TabPage tabAdvanced = new TabPage("Advanced");
            
            GroupBox grpServerAuth = new GroupBox { Text = "Server Authentication", Location = new Point(20, 15), Size = new Size(460, 70) };
            chkAdmin = new CheckBox { Text = "Connect to console / administrative session", Location = new Point(15, 30), Size = new Size(350, 25) };
            grpServerAuth.Controls.Add(chkAdmin);

            grpGateway = new GroupBox { Text = "RD Gateway Server Settings", Location = new Point(20, 95), Size = new Size(460, 310), Enabled = false };
            
            lblSelectedGatewayServer = new Label { Text = "Selected Server: (None)", Location = new Point(15, 25), AutoSize = true, Font = new Font(this.Font, FontStyle.Bold) };
            
            rdoGwAuto = new RadioButton { Text = "Automatically detect RD Gateway server settings", Location = new Point(15, 55), Size = new Size(400, 20) };
            rdoGwUse = new RadioButton { Text = "Use these RD Gateway server settings:", Location = new Point(15, 80), Size = new Size(400, 20) };
            
            Label lblGatewayHost = new Label { Text = "Server name:", Location = new Point(35, 105), AutoSize = true };
            txtGateway = new TextBox { Location = new Point(125, 102), Size = new Size(200, 25) };
            
            Label lblGwLogon = new Label { Text = "Logon method:", Location = new Point(35, 135), AutoSize = true };
            cmbGwLogon = new ComboBox { Location = new Point(125, 132), Size = new Size(200, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbGwLogon.Items.AddRange(new string[] { "Ask for password (NTLM)", "Smart card", "Allow me to select later" });
            cmbGwLogon.SelectedIndex = 2; // Default to 'Allow me to select later'
            
            chkGwBypass = new CheckBox { Text = "Bypass RD Gateway server for local addresses", Location = new Point(35, 165), Size = new Size(400, 20) };
            
            rdoGwNone = new RadioButton { Text = "Do not use an RD Gateway server", Location = new Point(15, 195), Size = new Size(400, 20), Checked = true };

            Label lblGwLogonSet = new Label { Text = "Logon settings", Location = new Point(15, 225), AutoSize = true, Font = new Font(this.Font, FontStyle.Bold) };
            
            chkGwReuseCreds = new CheckBox { Text = "Use my RD Gateway credentials for the remote computer", Location = new Point(15, 245), Size = new Size(400, 20) };

            Label lblGatewayNote = new Label { Text = "Note: Gateway settings are automatically saved as you configure them.", Location = new Point(15, 275), AutoSize = true, ForeColor = Color.DimGray };

            EventHandler gwChanged = (s, e) => {
                if (isUpdatingGateway) return;
                if (lstServers.SelectedIndex != -1) {
                    string srv = lstServers.SelectedItem.ToString();
                    
                    int usage = 0;
                    if (rdoGwAuto.Checked) usage = 2;
                    else if (rdoGwUse.Checked) usage = 1;
                    
                    string val = string.Format("{0}|{1}|{2}|{3}|{4}", 
                        usage, 
                        txtGateway.Text.Trim(), 
                        cmbGwLogon.SelectedIndex, 
                        chkGwBypass.Checked ? "1" : "0", 
                        chkGwReuseCreds.Checked ? "1" : "0");
                    dictGateways[srv] = val;

                    UpdateGatewayUIState();
                    
                    // Auto-save unless it's just a keystroke in the text box
                    if (s != txtGateway) SaveIni();
                }
            };

            rdoGwAuto.CheckedChanged += gwChanged;
            rdoGwUse.CheckedChanged += gwChanged;
            rdoGwNone.CheckedChanged += gwChanged;
            txtGateway.TextChanged += gwChanged;
            txtGateway.Leave += (s, e) => SaveIni(); // Save when focus leaves the textbox
            cmbGwLogon.SelectedIndexChanged += gwChanged;
            chkGwBypass.CheckedChanged += gwChanged;
            chkGwReuseCreds.CheckedChanged += gwChanged;

            grpGateway.Controls.AddRange(new Control[] { 
                lblSelectedGatewayServer, rdoGwAuto, rdoGwUse, 
                lblGatewayHost, txtGateway, lblGwLogon, cmbGwLogon, chkGwBypass, 
                rdoGwNone, lblGwLogonSet, chkGwReuseCreds, lblGatewayNote 
            });
            
            tabAdvanced.Controls.AddRange(new Control[] { grpServerAuth, grpGateway });

            // Add tabs to control
            tabMain.TabPages.AddRange(new TabPage[] { tabGeneral, tabDisplay, tabResources, tabAdvanced });
            tabMain.SelectedIndexChanged += (s, e) => SaveIni(); // Auto-save when switching tabs

            // --- BOTTOM BAR ---
            lnkGithub = new LinkLabel { Text = "GitHub: joshdwight101", Location = new Point(15, 555), AutoSize = true };
            lnkGithub.LinkClicked += (s, e) => Process.Start("https://github.com/joshdwight101");

            btnConnect = new Button { Text = "Connect", Location = new Point(310, 545), Size = new Size(105, 35), BackColor = Color.LightSkyBlue };
            btnConnect.Click += BtnConnect_Click;

            btnClose = new Button { Text = "Close", Location = new Point(425, 545), Size = new Size(105, 35) };
            btnClose.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[] { tabMain, lnkGithub, btnConnect, btnClose });
        }

        private void UpdateGatewayUIState()
        {
            bool useThese = rdoGwUse.Checked;
            if (txtGateway != null) txtGateway.Enabled = useThese;
            if (cmbGwLogon != null) cmbGwLogon.Enabled = useThese;
            if (chkGwBypass != null) chkGwBypass.Enabled = useThese;
        }

        private void LstServers_SelectedIndexChanged(object sender, EventArgs e)
        {
            isUpdatingGateway = true;
            if (lstServers.SelectedIndex != -1)
            {
                string srv = lstServers.SelectedItem.ToString();
                lblSelectedGatewayServer.Text = string.Format("Selected Server: {0}", srv);
                if (grpGateway != null) grpGateway.Enabled = true;
                
                if (dictGateways.ContainsKey(srv))
                {
                    string data = dictGateways[srv];
                    string[] parts = data.Split('|');
                    
                    if (parts.Length == 1) { // Legacy parsing for old configurations
                        rdoGwUse.Checked = true;
                        txtGateway.Text = parts[0];
                        cmbGwLogon.SelectedIndex = 2; // Allow to select later
                        chkGwBypass.Checked = true;
                        chkGwReuseCreds.Checked = false;
                    } else if (parts.Length >= 5) {
                        if (parts[0] == "0") rdoGwNone.Checked = true;
                        else if (parts[0] == "2") rdoGwAuto.Checked = true;
                        else rdoGwUse.Checked = true;
                        
                        txtGateway.Text = parts[1];
                        
                        int logonIdx;
                        if (int.TryParse(parts[2], out logonIdx) && logonIdx >= 0 && logonIdx < 3)
                            cmbGwLogon.SelectedIndex = logonIdx;
                        else
                            cmbGwLogon.SelectedIndex = 2;
                            
                        chkGwBypass.Checked = (parts[3] == "1");
                        chkGwReuseCreds.Checked = (parts[4] == "1");
                    }
                }
                else
                {
                    rdoGwNone.Checked = true;
                    txtGateway.Text = "";
                    cmbGwLogon.SelectedIndex = 2;
                    chkGwBypass.Checked = true;
                    chkGwReuseCreds.Checked = false;
                }
                UpdateGatewayUIState();
            }
            else
            {
                lblSelectedGatewayServer.Text = "Selected Server: (None)";
                if (grpGateway != null) grpGateway.Enabled = false;
                rdoGwNone.Checked = true;
                txtGateway.Text = "";
                cmbGwLogon.SelectedIndex = 2;
                chkGwBypass.Checked = true;
                chkGwReuseCreds.Checked = false;
                UpdateGatewayUIState();
            }
            isUpdatingGateway = false;
            SaveIni(); // Auto-save when switching servers
        }

        private void LoadIni()
        {
            if (!File.Exists(iniPath))
            {
                File.WriteAllLines(iniPath, new string[] { 
                    "[Servers]", 
                    "DC01.contoso.local", 
                    "FS01.contoso.local",
                    "",
                    "[Gateways]",
                    "; Format: Server=UsageMethod|Hostname|LogonIndex|BypassLocal|ReuseCreds",
                    "DC01.contoso.local=1|gw.contoso.local|2|1|0"
                });
            }

            lstServers.Items.Clear();
            dictGateways.Clear();
            var lines = File.ReadAllLines(iniPath);
            string currentSection = "";

            foreach (var line in lines)
            {
                string trimmed = line.Trim();
                if (trimmed.StartsWith("[") && trimmed.EndsWith("]")) 
                { 
                    currentSection = trimmed; 
                    continue; 
                }
                if (string.IsNullOrEmpty(trimmed) || trimmed.StartsWith(";")) 
                { 
                    continue; 
                }

                if (currentSection.Equals("[Servers]", StringComparison.OrdinalIgnoreCase))
                {
                    lstServers.Items.Add(trimmed);
                }
                else if (currentSection.Equals("[Gateways]", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = trimmed.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        dictGateways[parts[0].Trim()] = parts[1].Trim();
                    }
                }
            }
        }

        private void SaveIni()
        {
            List<string> lines = new List<string>();
            lines.Add("[Servers]");
            foreach (var item in lstServers.Items)
            {
                lines.Add(item.ToString());
            }
            
            lines.Add("");
            lines.Add("[Gateways]");
            foreach (var item in lstServers.Items)
            {
                string srv = item.ToString();
                if (dictGateways.ContainsKey(srv) && !string.IsNullOrEmpty(dictGateways[srv]))
                {
                    lines.Add(string.Format("{0}={1}", srv, dictGateways[srv]));
                }
            }
            
            File.WriteAllLines(iniPath, lines.ToArray());
        }

        private void BtnSaveIni_Click(object sender, EventArgs e)
        {
            SaveIni();
            MessageBox.Show("Server list and Gateway settings saved successfully to INI file.", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnRefreshIni_Click(object sender, EventArgs e)
        {
            LoadIni();
        }

        private void BtnAddServer_Click(object sender, EventArgs e)
        {
            string newServer = txtNewServer.Text.Trim();
            if (!string.IsNullOrEmpty(newServer) && !lstServers.Items.Contains(newServer))
            {
                lstServers.Items.Add(newServer);
                txtNewServer.Clear();
                lstServers.SelectedItem = newServer; // Auto-select to configure gateway
                SaveIni(); // Auto-save when adding
            }
        }

        private void BtnRemoveServer_Click(object sender, EventArgs e)
        {
            if (lstServers.SelectedIndex != -1)
            {
                string srv = lstServers.SelectedItem.ToString();
                lstServers.Items.RemoveAt(lstServers.SelectedIndex);
                if (dictGateways.ContainsKey(srv))
                {
                    dictGateways.Remove(srv);
                }
                SaveIni(); // Auto-save when removing
            }
        }

        private void RunCmdKey(string arguments)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("cmdkey.exe", arguments);
                psi.CreateNoWindow = true;
                psi.UseShellExecute = false;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                using (Process p = Process.Start(psi))
                {
                    p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error managing credentials: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSaveCreds_Click(object sender, EventArgs e)
        {
            string user = txtUsername.Text.Trim();
            string pass = txtPassword.Text;

            if (string.IsNullOrEmpty(user) || string.IsNullOrEmpty(pass))
            {
                MessageBox.Show("Please enter both username and password.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (lstServers.Items.Count == 0)
            {
                MessageBox.Show("No servers in the list to save credentials for.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            this.Cursor = Cursors.WaitCursor;
            foreach (var item in lstServers.Items)
            {
                string srv = item.ToString();
                RunCmdKey(string.Format("/generic:\"TERMSRV/{0}\" /user:\"{1}\" /pass:\"{2}\"", srv, user, pass));
            }
            this.Cursor = Cursors.Default;

            MessageBox.Show(string.Format("Credentials successfully deployed for {0} server(s).\nYou will now automatically log in when connecting via RDP.", lstServers.Items.Count), 
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            txtPassword.Clear();
        }

        private void BtnClearCreds_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to clear the saved credentials for ALL servers in the list?", "Confirm Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                this.Cursor = Cursors.WaitCursor;
                foreach (var item in lstServers.Items)
                {
                    string srv = item.ToString();
                    RunCmdKey(string.Format("/delete:\"TERMSRV/{0}\"", srv));
                }
                this.Cursor = Cursors.Default;

                MessageBox.Show("Credentials cleared successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnConnect_Click(object sender, EventArgs e)
        {
            if (lstServers.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a server from the list first.", "Select Server", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string targetServer = lstServers.SelectedItem.ToString();
            
            try
            {
                string rdpPath = Path.Combine(Path.GetTempPath(), "TermServPlus_Temp.rdp");
                List<string> rdpLines = new List<string>();
                
                rdpLines.Add(string.Format("full address:s:{0}", targetServer));
                
                // --- Display Settings ---
                if (cmbResolution.SelectedItem != null && cmbResolution.SelectedItem.ToString() == "Full Screen")
                {
                    rdpLines.Add("screen mode id:i:2");
                }
                else
                {
                    rdpLines.Add("screen mode id:i:1");
                    string res = cmbResolution.SelectedItem != null ? cmbResolution.SelectedItem.ToString() : "1024x768";
                    string[] parts = res.Split('x');
                    if (parts.Length == 2)
                    {
                        rdpLines.Add(string.Format("desktopwidth:i:{0}", parts[0]));
                        rdpLines.Add(string.Format("desktopheight:i:{0}", parts[1]));
                    }
                }

                if (chkMultiMon.Checked) rdpLines.Add("use multimon:i:1");

                string color = cmbColors.SelectedItem != null ? cmbColors.SelectedItem.ToString() : "True Color (24 bit)";
                int colorDepth = 24;
                if (color.Contains("15")) colorDepth = 15;
                else if (color.Contains("16")) colorDepth = 16;
                else if (color.Contains("32")) colorDepth = 32;
                rdpLines.Add(string.Format("session bpp:i:{0}", colorDepth));

                // --- Local Resources Settings ---
                // Audio
                int audioMode = 0; // default: bring to computer
                if (cmbAudioPlayback.SelectedIndex == 1) audioMode = 1; // Play on remote
                else if (cmbAudioPlayback.SelectedIndex == 2) audioMode = 2; // Do not play
                rdpLines.Add(string.Format("audiomode:i:{0}", audioMode));
                rdpLines.Add(string.Format("audiocapturemode:i:{0}", chkAudioRecord.Checked ? "1" : "0"));
                
                // Keyboard
                int keybHook = 2; // default: full screen
                if (cmbKeyboard.SelectedIndex == 0) keybHook = 0; // On this computer
                else if (cmbKeyboard.SelectedIndex == 1) keybHook = 1; // On remote
                rdpLines.Add(string.Format("keyboardhook:i:{0}", keybHook));

                // Devices
                rdpLines.Add(string.Format("redirectprinters:i:{0}", chkPrinters.Checked ? "1" : "0"));
                rdpLines.Add(string.Format("redirectclipboard:i:{0}", chkClipboard.Checked ? "1" : "0"));

                // --- Advanced Settings ---
                if (chkAdmin.Checked) rdpLines.Add("administrative session:i:1");

                // Gateway
                if (dictGateways.ContainsKey(targetServer) && !string.IsNullOrEmpty(dictGateways[targetServer]))
                {
                    string data = dictGateways[targetServer];
                    string[] parts = data.Split('|');
                    int usage = 0;
                    string host = "";
                    int logonIdx = 2;
                    bool bypass = true;
                    bool reuse = false;

                    if (parts.Length == 1) { // Legacy parsing
                        usage = string.IsNullOrEmpty(parts[0]) ? 0 : 1;
                        host = parts[0];
                    } else if (parts.Length >= 5) {
                        int.TryParse(parts[0], out usage);
                        host = parts[1];
                        int.TryParse(parts[2], out logonIdx);
                        bypass = parts[3] == "1";
                        reuse = parts[4] == "1";
                    }

                    rdpLines.Add(string.Format("gatewayusagemethod:i:{0}", usage));
                    
                    if (usage == 1 || usage == 2) {
                        rdpLines.Add(string.Format("gatewayhostname:s:{0}", host));
                        
                        // Map our combobox index to RDP gatewaycredentialssource
                        int credSource = 4; // Select later
                        if (logonIdx == 0) credSource = 0; // Password
                        else if (logonIdx == 1) credSource = 1; // Smart card
                        rdpLines.Add(string.Format("gatewaycredentialssource:i:{0}", credSource));
                        
                        rdpLines.Add(string.Format("gatewaybypasslocaladdress:i:{0}", bypass ? "1" : "0"));
                        rdpLines.Add(string.Format("promptcredentialonce:i:{0}", reuse ? "1" : "0"));
                    }
                }
                else
                {
                    rdpLines.Add("gatewayusagemethod:i:0"); // 0 = do not use
                }

                File.WriteAllLines(rdpPath, rdpLines.ToArray());

                ProcessStartInfo psi = new ProcessStartInfo("mstsc.exe", string.Format("\"{0}\"", rdpPath));
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to launch RDP: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
'@

# Add necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Enable visual styles natively in PowerShell (safely ignoring if already set in this session)
try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }

# Generate a unique namespace ID to completely prevent "Type already exists" errors when testing
$uniqueId = [guid]::NewGuid().ToString("N")
$compilableCode = $csharpCode.Replace("namespace TermServPlusApp", "namespace TermServPlusApp_$uniqueId")

# Compile the C# code into memory
try {
    Add-Type -TypeDefinition $compilableCode -ReferencedAssemblies "System.Windows.Forms", "System.Drawing" -ErrorAction Stop
}
catch {
    Write-Host "Failed to compile C# GUI. Error: $_" -ForegroundColor Red
    Exit
}

# Safely instantiate and show the form without conflicting with Application.Run() lifecycle limits
$formType = "TermServPlusApp_${uniqueId}.TermServForm"
$form = New-Object $formType -ArgumentList $scriptRoot
$form.ShowDialog() | Out-Null
$form.Dispose()