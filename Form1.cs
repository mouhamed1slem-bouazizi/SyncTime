using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using ExcelDataReader;
using Renci.SshNet;
using System.IO;
using System.Net.NetworkInformation;
using ClosedXML.Excel;
using System.Diagnostics;
using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Serialization;

namespace SyncTime
{
    [Serializable]
    public class ClientInfo
    {
        public string Name { get; set; }
        public string IP { get; set; }
        public string MacAddress { get; set; }
        public string ActualMacAddress { get; set; }
        public string Monitor { get; set; }
        public string Type { get; set; }
        public string Group { get; set; }
        public string Level { get; set; }
        public string Zone { get; set; }
        public DateTime? LastCheckTime { get; set; }
        public TimeSpan? TimeDrift { get; set; }
        public bool IsDriftCritical { get; set; }
        [XmlIgnore] // Ignore this property during serialization
        public List<TimeDriftRecord> DriftHistory { get; set; } = new List<TimeDriftRecord>();
    }

    public class TimeDriftRecord
    {
        public DateTime CheckTime { get; set; }
        public TimeSpan Drift { get; set; }
        public string Status { get; set; }
        public string DisplayTime { get; set; }
    }
    public partial class Form1 : Form
    {
        private const string EXCEL_FILE_PATH = @"C:\IP\Dammam_Inventory.xlsx";
        private const string EXCEL_SHEET_NAME = "FIDS_Inv";

        // Fields declaration
        private List<ClientInfo> clients;
        private DataGridView gridView;
        private Button syncButton;
        private Button executeCodeButton;
        private Button exportButton;
        private Label statusLabel;
        private Button statisticsButton;
        private Button batchOperationsButton;
        private Button historyButton;
        private List<ConnectionLog> connectionHistory;
        private TableLayoutPanel mainLayout;
        private DataGridViewRow selectedRow;
        private TextBox searchBox;
        private ComboBox filterColumn;
        private ComboBox monitorFilter;
        private ComboBox typeFilter;
        private ComboBox groupFilter;
        private ComboBox levelFilter;
        private ComboBox zoneFilter;
        private Panel searchAndFilterPanel;
        private Button checkTimeButton;
        private Button checkMacButton;
        private const string CACHE_FILE_PATH = @"C:\IP\client_cache.xml";
        private Button loadClientsButton;
        private Button modifyButton;

        public Form1()
        {
            try
            {
                // Initialize collections first
                clients = new List<ClientInfo>();
                connectionHistory = new List<ConnectionLog>();

                InitializeComponent();

                // Create menu strip first (it should be at the top)
                AddCacheManagementMenu();

                // Initialize the main UI components
                InitializeUI();

                // Add all additional controls after main UI is set up
                mainLayout.SuspendLayout();
                AddLoadButton();  // Add load button first since it goes at the top
                AddSearchAndFilter();
                AddFilterControls();
                AddCheckButtons();
                mainLayout.ResumeLayout();

                // Initialize context menu for drift history
                InitializeTimeDriftTracking();

                // Load cached data last
                LoadCachedClients();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing application: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}",
                    "Initialization Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private class ClientStatistics
        {
            public int TotalClients { get; set; }
            public int OnlineClients { get; set; }
            public int OfflineClients { get; set; }
            public int SynchronizedClients { get; set; }
            public int MatchedMacs { get; set; }
            public int MismatchedMacs { get; set; }
            public Dictionary<string, int> ClientsBySystem { get; set; }
            public Dictionary<string, int> ClientsByStatus { get; set; }
            public DateTime LastRefresh { get; set; }
            public Dictionary<string, int> DriftStats { get; set; }
            public List<(string Name, TimeSpan Drift)> CriticalDriftClients { get; set; }
        }

        public class ConnectionLog
        {
            public DateTime Timestamp { get; set; }
            public string ClientName { get; set; }
            public string IP { get; set; }
            public string Action { get; set; }
            public string Status { get; set; }
            public string Details { get; set; }
        }

        private void InitializeUI()
        {
            this.Size = new Size(900, 600);
            this.MinimumSize = new Size(800, 500);
            this.Text = "FIDS Client Time Sync";

            // Create main layout panel
            mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10),
                AutoSize = true,
                Margin = new Padding(25, 25, 25, 25)
            };

            // Configure row styles
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // Status label
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Grid
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));  // Buttons

            this.Controls.Add(mainLayout);

            // Status Label
            statusLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.Blue,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 5, 0, 5),
                Margin = new Padding(0, 14, 0, 0)
            };
            mainLayout.Controls.Add(statusLabel, 0, 0);

            // Grid View Panel
            Panel gridPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 5, 0, 5)
            };
            mainLayout.Controls.Add(gridPanel, 0, 1);

            modifyButton = new Button
            {
                Text = "Modify",
                Size = new Size(150, 30),
                Enabled = false,  // Will be enabled when a row is selected
                Anchor = AnchorStyles.None
            };
            modifyButton.Click += ModifyButton_Click;

            // Grid View
            gridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                BackgroundColor = SystemColors.Control,
                BorderStyle = BorderStyle.Fixed3D,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };

            // Initialize columns
            gridView.Columns.Add("Name", "Client Name");
            gridView.Columns.Add("IP", "IP Address");
            gridView.Columns.Add("MAC", "MAC Address");
            gridView.Columns.Add("ActualMAC", "Actual MAC");
            gridView.Columns.Add("Status", "Status");
            gridView.Columns.Add("CurrentTime", "Current Time");
            gridView.Columns.Add("LastSync", "Last Sync Time");
            gridView.Columns.Add("Monitor", "Monitor");
            gridView.Columns.Add("Type", "Type");
            gridView.Columns.Add("Group", "Group");
            gridView.Columns.Add("Level", "Level");
            gridView.Columns.Add("Zone", "Zone");

            // Set column weights
            gridView.Columns[0].FillWeight = 15;  // Name
            gridView.Columns[1].FillWeight = 15;  // IP
            gridView.Columns[2].FillWeight = 15;  // MAC
            gridView.Columns[3].FillWeight = 15;  // Actual MAC
            gridView.Columns[4].FillWeight = 10;  // Status
            gridView.Columns[5].FillWeight = 15;  // Current Time
            gridView.Columns[6].FillWeight = 15;  // Last Sync
            gridView.Columns[7].FillWeight = 10;  // Monitor
            gridView.Columns[8].FillWeight = 10;  // Type
            gridView.Columns[9].FillWeight = 10;  // Group
            gridView.Columns[10].FillWeight = 10; // Level
            gridView.Columns[11].FillWeight = 10; // Zone

            gridPanel.Controls.Add(gridView);

            // Add MAC address copy functionality
            gridView.CellClick += (sender, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex == 3)
                {
                    var macAddress = gridView.Rows[e.RowIndex].Cells[3].Value?.ToString();
                    if (!string.IsNullOrEmpty(macAddress) && macAddress != "-")
                    {
                        try
                        {
                            Clipboard.SetText(macAddress);

                            var originalColor = gridView.Rows[e.RowIndex].Cells[3].Style.BackColor;
                            gridView.Rows[e.RowIndex].Cells[3].Style.BackColor = Color.Yellow;

                            var tooltip = new ToolTip();
                            var relativeMousePos = gridView.PointToClient(Cursor.Position);
                            tooltip.Show("MAC Address Copied!", gridView, relativeMousePos.X + 15, relativeMousePos.Y, 1000);

                            Task.Delay(200).ContinueWith(t =>
                            {
                                if (gridView.InvokeRequired)
                                {
                                    gridView.Invoke(new Action(() =>
                                    {
                                        gridView.Rows[e.RowIndex].Cells[3].Style.BackColor = originalColor;
                                    }));
                                }
                                else
                                {
                                    gridView.Rows[e.RowIndex].Cells[3].Style.BackColor = originalColor;
                                }
                            });
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Failed to copy MAC address: {ex.Message}", "Copy Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            };

            // Add cell formatting for MAC address column
            gridView.CellFormatting += (sender, e) =>
            {
                if (e.ColumnIndex == 3 && e.RowIndex >= 0)
                {
                    var cell = gridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    var value = cell.Value?.ToString();
                    if (!string.IsNullOrEmpty(value) && value != "-")
                    {
                        var currentFont = cell.Style.Font ?? gridView.DefaultCellStyle.Font;
                        cell.Style.Font = new Font(currentFont, FontStyle.Underline);
                        if (cell.Tag == null)
                        {
                            cell.Tag = Cursors.Hand;
                        }
                    }
                }
            };

            // Add cursor change handlers
            gridView.CellMouseEnter += (sender, e) =>
            {
                if (e.ColumnIndex == 3 && e.RowIndex >= 0)
                {
                    var cell = gridView.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    if (cell.Tag is Cursor cursor)
                    {
                        gridView.Cursor = cursor;
                    }
                }
            };

            gridView.CellMouseLeave += (sender, e) =>
            {
                gridView.Cursor = Cursors.Default;
            };

            // Add selection change handler for the grid
            gridView.CellClick += GridView_CellClick;

            // Button Panel
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill
            };
            mainLayout.Controls.Add(buttonPanel, 0, 2);

            // Create a TableLayoutPanel for the buttons
            TableLayoutPanel buttonsLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 1
            };
            buttonPanel.Controls.Add(buttonsLayout);

            // Sync Button
            syncButton = new Button
            {
                Text = "Synchronize Time",
                Size = new Size(150, 30),
                Enabled = false,
                Anchor = AnchorStyles.None
            };
            syncButton.Click += SyncButton_Click;

            // Execute Code Button
            executeCodeButton = new Button
            {
                Text = "Execute Code",
                Size = new Size(150, 30),
                Enabled = false,
                Anchor = AnchorStyles.None
            };
            executeCodeButton.Click += ExecuteCodeButton_Click;

            // Create export button
            exportButton = new Button
            {
                Text = "Export Data",
                Size = new Size(150, 30),
                Enabled = true,
                Anchor = AnchorStyles.None
            };
            exportButton.Click += ExportButton_Click;

            // Statistics button
            statisticsButton = new Button
            {
                Text = "Statistics",
                Size = new Size(150, 30),
                Enabled = true,
                Anchor = AnchorStyles.None
            };
            statisticsButton.Click += StatisticsButton_Click;

            // Batch operations button
            batchOperationsButton = new Button
            {
                Text = "Batch Operations",
                Size = new Size(150, 30),
                Enabled = false,
                Anchor = AnchorStyles.None
            };
            batchOperationsButton.Click += BatchOperationsButton_Click;

            // History button
            historyButton = new Button
            {
                Text = "Connection History",
                Size = new Size(150, 30),
                Enabled = true,
                Anchor = AnchorStyles.None
            };
            historyButton.Click += HistoryButton_Click;

            // Add buttons to the layout
            buttonsLayout.Controls.Add(syncButton, 0, 0);
            buttonsLayout.Controls.Add(executeCodeButton, 1, 0);
            buttonsLayout.Controls.Add(exportButton, 2, 0);
            buttonsLayout.Controls.Add(statisticsButton, 3, 0);
            buttonsLayout.Controls.Add(batchOperationsButton, 4, 0);
            buttonsLayout.Controls.Add(historyButton, 5, 0);
            buttonsLayout.Controls.Add(modifyButton, 6, 0);  // Adjust the column index as needed

            // Center the buttons in their cells
            for (int i = 0; i < 6; i++)
            {
                buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20f));
            }

            // Style the grid
            gridView.EnableHeadersVisualStyles = false;
            gridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 66, 91);
            gridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            gridView.ColumnHeadersDefaultCellStyle.Font = new Font(gridView.Font, FontStyle.Bold);
            gridView.ColumnHeadersHeight = 35;
            gridView.RowTemplate.Height = 30;
            gridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250);
            gridView.MultiSelect = true;
            gridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            gridView.SelectionChanged += GridView_SelectionChanged;

            // Set initial status
            statusLabel.Text = "Loading clients...";
        }

        private void GridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                selectedRow = gridView.Rows[e.RowIndex];
                executeCodeButton.Enabled = true;
            }
        }

        private async void ExecuteCodeButton_Click(object sender, EventArgs e)
        {
            if (selectedRow == null) return;

            string clientName = selectedRow.Cells[0].Value.ToString();
            string clientIP = selectedRow.Cells[1].Value.ToString();

            // Create and configure the input dialog
            Form promptForm = new Form()
            {
                Width = 800,
                Height = 600,
                FormBorderStyle = FormBorderStyle.Sizable,
                Text = $"Execute Code on {clientName} ({clientIP})",
                StartPosition = FormStartPosition.CenterParent,
                MinimumSize = new Size(600, 400)
            };

            TableLayoutPanel mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };

            // Command input box
            TextBox commandBox = new TextBox()
            {
                Multiline = true,
                Height = 80,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical
            };

            // Results box
            TextBox resultsBox = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Both,
                BackColor = Color.Black,
                ForeColor = Color.LightGreen,
                Font = new Font("Consolas", 10)
            };

            // Buttons panel
            TableLayoutPanel buttonPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1
            };

            Button executeButton = new Button()
            {
                Text = "Execute",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.None
            };

            Button clearButton = new Button()
            {
                Text = "Clear Results",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.None
            };

            Button closeButton = new Button()
            {
                Text = "Close Session",
                Width = 100,
                Height = 30,
                Anchor = AnchorStyles.None
            };

            // Configure layout
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));  // Command box
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // Results box
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));   // Buttons

            buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
            buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));

            // Add controls
            mainPanel.Controls.Add(commandBox, 0, 0);
            mainPanel.Controls.Add(resultsBox, 0, 1);
            mainPanel.Controls.Add(buttonPanel, 0, 2);

            buttonPanel.Controls.Add(executeButton, 0, 0);
            buttonPanel.Controls.Add(clearButton, 1, 0);
            buttonPanel.Controls.Add(closeButton, 2, 0);

            promptForm.Controls.Add(mainPanel);

            // Create SSH client
            SshClient sshClient = null;
            bool isConnected = false;

            // Connect to client
            try
            {
                LogConnection(clientName, clientIP, "Execute Code", "Connecting");
                UpdateGridRow(clientName, clientIP, "Connecting...");
                var ping = new Ping();
                var reply = await ping.SendPingAsync(clientIP, 1000);

                if (reply.Status != IPStatus.Success)
                {
                    LogConnection(clientName, clientIP, "Execute Code", "Unreachable");
                    UpdateGridRow(clientName, clientIP, "Unreachable");
                    MessageBox.Show("Client is unreachable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                sshClient = new SshClient(clientIP, "root", "123456");
                await Task.Run(() => sshClient.Connect());
                isConnected = true;
                UpdateGridRow(clientName, clientIP, "Connected");
                LogConnection(clientName, clientIP, "Execute Code", "Connected");

                resultsBox.AppendText("Connected to " + clientIP + "\r\n");

            }
            catch (Exception ex)
            {
                LogConnection(clientName, clientIP, "Execute Code", "Failed", ex.Message);
                MessageBox.Show($"Failed to connect: {ex.Message}", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateGridRow(clientName, clientIP, "Connection Failed");
                return;
            }

            // Handle execute button click
            executeButton.Click += async (s, ev) =>
            {
                if (!isConnected || sshClient == null) return;

                string command = commandBox.Text.Trim();
                if (string.IsNullOrEmpty(command)) return;

                try
                {
                    resultsBox.AppendText($"\r\n> {command}\r\n");
                    var result = await Task.Run(() => sshClient.RunCommand(command).Result);
                    resultsBox.AppendText($"{result}\r\n");
                    commandBox.Clear();
                }
                catch (Exception ex)
                {
                    resultsBox.AppendText($"Error: {ex.Message}\r\n");
                }
            };

            // Handle clear button click
            clearButton.Click += (s, ev) =>
            {
                resultsBox.Clear();
            };

            // Handle close button click
            closeButton.Click += (s, ev) =>
            {
                promptForm.Close();
            };

            // Handle form closing
            promptForm.FormClosing += (s, ev) =>
            {
                if (isConnected && sshClient != null)
                {
                    sshClient.Disconnect();
                    sshClient.Dispose();
                }
                UpdateGridRow(clientName, clientIP, "Disconnected");
            };

            // Enable/disable buttons based on connection status
            executeButton.Enabled = isConnected;

            // Show the form
            promptForm.ShowDialog();
        }
        /*private async void FileSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filePath = Path.Combine(EXCEL_FILES_DIRECTORY, "FIDS_Inventory.xlsx"); // Your Excel file name

            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Excel file not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    statusLabel.Text = "Excel file not found";
                    statusLabel.ForeColor = Color.Red;
                    return;
                }

                statusLabel.Text = "Loading clients...";
                statusLabel.ForeColor = Color.Blue;

                await Task.Run(() => LoadIPsFromExcel(filePath));
                UpdateFilterOptions();

                if (clients.Any())
                {
                    syncButton.Enabled = true;
                    await RefreshClientTimes();
                    statusLabel.Text = $"Loaded {clients.Count} clients";
                    statusLabel.ForeColor = Color.Green;
                }
                else
                {
                    statusLabel.Text = "No clients found in file";
                    statusLabel.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Error loading clients";
                statusLabel.ForeColor = Color.Red;
            }
        }*/

        /*private async Task LoadAllClients()
        {
            int totalFiles = excelFiles.Count - 1; // Excluding ALL_CLIENTS_KEY
            int loadedFiles = 0;
            int totalClients = 0;
            int successfullyLoaded = 0;
            statusLabel.Text = "Loading all clients...";
            statusLabel.ForeColor = Color.Blue;

            clients.Clear();
            gridView.Invoke(new Action(() => gridView.Rows.Clear()));

            foreach (var file in excelFiles)
            {
                if (file.Key == ALL_CLIENTS_KEY) continue;
                loadedFiles++;

                string filePath = Path.Combine(EXCEL_FILES_DIRECTORY, file.Value);
                if (File.Exists(filePath))
                {
                    try
                    {
                        int initialCount = clients.Count;
                        await Task.Run(() => LoadIPsFromExcel(filePath));
                        int newClientsAdded = clients.Count - initialCount;
                        totalClients += newClientsAdded;
                        successfullyLoaded++;

                        statusLabel.Text = $"Loading files... ({loadedFiles}/{totalFiles}) - {totalClients} clients loaded";
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error loading {file.Key}: {ex.Message}", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            // Show final summary
            if (successfullyLoaded > 0)
            {
                statusLabel.Text = $"Loaded {totalClients} total clients from {successfullyLoaded} systems";
                statusLabel.ForeColor = Color.Green;
            }
            else
            {
                statusLabel.Text = "No clients were loaded";
                statusLabel.ForeColor = Color.Red;
            }
        }*/

        // Add this class to track loading statistics
        public class LoadingStats
        {
            public int TotalFiles { get; set; }
            public int LoadedFiles { get; set; }
            public int TotalClients { get; set; }
            public int SuccessfulClients { get; set; }
            public int FailedClients { get; set; }
            public Dictionary<string, int> ClientsPerSystem { get; set; } = new Dictionary<string, int>();
        }

        private void LoadIPsFromExcel(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(EXCEL_SHEET_NAME);
                var rows = worksheet.RowsUsed().Skip(1); // Skip header row

                clients.Clear();
                if (gridView.InvokeRequired)
                {
                    gridView.Invoke(new Action(() => gridView.Rows.Clear()));
                }
                else
                {
                    gridView.Rows.Clear();
                }

                foreach (var row in rows)
                {
                    try
                    {
                        string name = row.Cell("A").GetString().Trim();
                        string ip = row.Cell("B").GetString().Trim();
                        string mac = row.Cell("C").GetString().Trim().ToUpper();
                        string monitor = row.Cell("D").GetString().Trim();
                        string type = row.Cell("E").GetString().Trim();
                        string group = row.Cell("F").GetString().Trim();
                        string level = row.Cell("G").GetString().Trim();
                        string zone = row.Cell("H").GetString().Trim();

                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(ip))
                        {
                            var client = new ClientInfo
                            {
                                Name = name,
                                IP = ip,
                                MacAddress = mac,
                                ActualMacAddress = "",
                                Monitor = monitor,
                                Type = type,
                                Group = group,
                                Level = level,
                                Zone = zone
                            };

                            clients.Add(client);

                            if (gridView.InvokeRequired)
                            {
                                gridView.Invoke(new Action(() =>
                                {
                                    AddRowToGrid(client);
                                }));
                            }
                            else
                            {
                                AddRowToGrid(client);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error but continue processing other rows
                        Debug.WriteLine($"Error processing row: {ex.Message}");
                        continue;
                    }
                }

                // Update filters after loading data
                if (gridView.InvokeRequired)
                {
                    gridView.Invoke(new Action(() => UpdateFilterOptions()));
                }
                else
                {
                    UpdateFilterOptions();
                }
            }
        }

        private void AddRowToGrid(ClientInfo client)
        {
            int rowIndex = gridView.Rows.Add(
                client.Name,           // Column 0: Client Name
                client.IP,            // Column 1: IP Address
                client.MacAddress,    // Column 2: MAC Address
                "-",                  // Column 3: Actual MAC
                "Not checked",        // Column 4: Status
                "-",                  // Column 5: Current Time
                "-",                  // Column 6: Last Sync Time
                client.Monitor,       // Column 7: Monitor
                client.Type,          // Column 8: Type
                client.Group,         // Column 9: Group
                client.Level,         // Column 10: Level
                client.Zone           // Column 11: Zone
            );

            // Apply any initial formatting
            var row = gridView.Rows[rowIndex];

            // Set tooltip for monitor and type columns
            row.Cells[7].ToolTipText = client.Monitor;
            row.Cells[8].ToolTipText = client.Type;

            // Set cell styles
            foreach (DataGridViewCell cell in row.Cells)
            {
                if (string.IsNullOrEmpty(cell.Value?.ToString()) || cell.Value.ToString() == "-")
                {
                    cell.Style.BackColor = Color.LightGray;
                    cell.Value = "-";
                }
            }
        }

        private async void SyncButton_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Are you sure you want to synchronize time on all clients?",
                "Confirm Synchronization",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                syncButton.Enabled = false;
                statusLabel.Text = "Synchronizing time...";
                statusLabel.ForeColor = Color.Blue;

                try
                {
                    await SynchronizeTime();
                    statusLabel.Text = "Time synchronization completed";
                    statusLabel.ForeColor = Color.Green;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error during synchronization: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    statusLabel.Text = "Synchronization failed";
                    statusLabel.ForeColor = Color.Red;
                }
                finally
                {
                    syncButton.Enabled = true;
                }
            }
        }

        // Modify the SynchronizeTime method to preserve MAC addresses
        private async Task SynchronizeTime()
        {
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            List<string> unreachableClients = new List<string>();

            await Parallel.ForEachAsync(clients, async (client, token) =>
            {
                try
                {
                    LogConnection(client.Name, client.IP, "Sync Time", "Starting");

                    var ping = new Ping();
                    var reply = await ping.SendPingAsync(client.IP, 1000);

                    if (reply.Status != IPStatus.Success)
                    {
                        LogConnection(client.Name, client.IP, "Sync Time", "Unreachable");
                        UpdateGridRow(client.Name, client.IP, "Unreachable");
                        unreachableClients.Add($"{client.Name} ({client.IP})");
                        return;
                    }

                    using (var sshClient = new SshClient(client.IP, "root", "123456"))
                    {
                        UpdateGridRow(client.Name, client.IP, "Connecting...");
                        await Task.Run(() => sshClient.Connect());

                        UpdateGridRow(client.Name, client.IP, "Setting time...");

                        // Get MAC address before setting time
                        var macCommand = "ip link show | grep -i 'link/ether' | awk '{print $2}' | head -n 1";
                        var actualMac = await Task.Run(() =>
                            sshClient.RunCommand(macCommand).Result.Trim().ToUpper());

                        // Set the time
                        string dateCommand = $"date -s \"{currentTime}\"";
                        await Task.Run(() => sshClient.RunCommand(dateCommand));
                        await Task.Run(() => sshClient.RunCommand("hwclock --systohc"));

                        // Verify time was set correctly
                        var timeOutput = await Task.Run(() =>
                            sshClient.RunCommand("date '+%Y-%m-%d %H:%M:%S'").Result.Trim());

                        client.ActualMacAddress = actualMac;

                        sshClient.Disconnect();
                        UpdateGridRowWithMac(client.Name, client.IP, client.MacAddress, actualMac, "Synchronized", timeOutput);
                        LogConnection(client.Name, client.IP, "Sync Time", "Synchronized", $"Time set to: {timeOutput}");
                    }
                }
                catch (Exception ex)
                {
                    LogConnection(client.Name, client.IP, "Sync Time", "Failed", ex.Message);
                    UpdateGridRow(client.Name, client.IP, "Connection Failed");
                    unreachableClients.Add($"{client.Name} ({client.IP})");
                }
            });

            if (unreachableClients.Any())
            {
                string message = "The following clients were unreachable or failed:\n\n" +
                                string.Join("\n", unreachableClients);
                MessageBox.Show(message, "Unreachable Clients", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void UpdateGridRow(string name, string ip, string status, string time = "-")
        {
            var client = clients.FirstOrDefault(c => c.Name == name && c.IP == ip);
            if (client != null)
            {
                UpdateGridRowWithMac(name, ip, client.MacAddress, client.ActualMacAddress, status, time);
                return;
            }

            if (gridView.InvokeRequired)
            {
                gridView.Invoke(new Action(() => UpdateGridRow(name, ip, status, time)));
                return;
            }

            foreach (DataGridViewRow row in gridView.Rows)
            {
                if (row.Cells[0].Value.ToString() == name && row.Cells[1].Value.ToString() == ip)
                {
                    // Preserve MAC addresses (cells 2 and 3)
                    string existingMac = row.Cells[2].Value?.ToString() ?? "-";
                    string existingActualMac = row.Cells[3].Value?.ToString() ?? "-";

                    // Update status in the correct cell (cell 4)
                    row.Cells[4].Value = status;

                    // Update time if provided (cell 5)
                    if (time != "-")
                    {
                        row.Cells[5].Value = time;
                        if (status == "Synchronized")
                        {
                            row.Cells[6].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }

                    // Color code the status cell
                    switch (status)
                    {
                        case "Synchronized":
                            row.Cells[4].Style.BackColor = Color.LightGreen;
                            break;
                        case "Unreachable":
                        case "Connection Failed":
                            row.Cells[4].Style.BackColor = Color.LightPink;
                            break;
                        case "Connecting...":
                        case "Setting time...":
                            row.Cells[4].Style.BackColor = Color.LightYellow;
                            break;
                        default:
                            row.Cells[4].Style.BackColor = Color.White;
                            break;
                    }
                    break;
                }
            }
        }

        private async Task RefreshClientTimes()
        {
            foreach (var client in clients)
            {
                try
                {
                    var ping = new Ping();
                    var reply = await ping.SendPingAsync(client.IP, 1000);

                    if (reply.Status != IPStatus.Success)
                    {
                        UpdateGridRow(client.Name, client.IP, "Unreachable");
                        continue;
                    }

                    using (var sshClient = new SshClient(client.IP, "root", "123456"))
                    {
                        await Task.Run(() => sshClient.Connect());

                        // Get current time and check drift
                        await CheckTimeDrift(client);

                        // Get MAC address
                        var macCommand = "ip link show | grep -i 'link/ether' | awk '{print $2}' | head -n 1";
                        var actualMac = await Task.Run(() =>
                            sshClient.RunCommand(macCommand).Result.Trim().ToUpper());

                        client.ActualMacAddress = actualMac;

                        sshClient.Disconnect();

                        string status = client.IsDriftCritical ? "Time Drift Critical" : "Connected";
                        UpdateGridRowWithMac(client.Name, client.IP, client.MacAddress, actualMac, status);
                    }
                }
                catch (Exception ex)
                {
                    LogConnection(client.Name, client.IP, "Refresh", "Failed", ex.Message);
                    UpdateGridRow(client.Name, client.IP, "Connection Failed");
                }
            }
        }
        private void UpdateGridRowWithMac(string name, string ip, string expectedMac, string actualMac, string status, string time = "-")
        {
            if (gridView.InvokeRequired)
            {
                gridView.Invoke(new Action(() => UpdateGridRowWithMac(name, ip, expectedMac, actualMac, status, time)));
                return;
            }

            foreach (DataGridViewRow row in gridView.Rows)
            {
                if (row.Cells[0].Value.ToString() == name && row.Cells[1].Value.ToString() == ip)
                {
                    // Update MAC addresses
                    row.Cells[2].Value = expectedMac ?? "-";
                    row.Cells[3].Value = actualMac ?? "-";

                    // Update status
                    row.Cells[4].Value = status;

                    // Update times
                    if (time != "-")
                    {
                        row.Cells[5].Value = time;
                        if (status == "Synchronized")
                        {
                            row.Cells[6].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }

                    // Handle MAC address highlighting
                    if (!string.IsNullOrEmpty(actualMac) && !string.IsNullOrEmpty(expectedMac))
                    {
                        if (actualMac != expectedMac)
                        {
                            row.Cells[3].Style.BackColor = Color.LightGreen;
                            row.Cells[3].Style.ForeColor = Color.Black;
                        }
                        else
                        {
                            row.Cells[3].Style.BackColor = Color.White;
                            row.Cells[3].Style.ForeColor = Color.Black;
                        }
                    }

                    // Update status cell color
                    switch (status)
                    {
                        case "Synchronized":
                            row.Cells[4].Style.BackColor = Color.LightGreen;
                            break;
                        case "Unreachable":
                        case "Connection Failed":
                            row.Cells[4].Style.BackColor = Color.LightPink;
                            break;
                        case "Connecting...":
                        case "Setting time...":
                            row.Cells[4].Style.BackColor = Color.LightYellow;
                            break;
                        default:
                            row.Cells[4].Style.BackColor = Color.White;
                            break;
                    }
                    break;
                }
            }
        }
        private void AddSearchAndFilter()
        {
            // Create panel for search controls
            searchAndFilterPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                Height = 40
            };

            // Configure column styles
            ((TableLayoutPanel)searchAndFilterPanel).ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            ((TableLayoutPanel)searchAndFilterPanel).ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80));
            ((TableLayoutPanel)searchAndFilterPanel).ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            // Initialize search box
            searchBox = new TextBox
            {
                PlaceholderText = "Search...",
                Width = 200,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(5)
            };

            // Initialize filter dropdown
            filterColumn = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 150,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(5)
            };

            filterColumn.Items.AddRange(new string[]
            {
        "All Columns",
        "Client Name",
        "IP Address",
        "MAC Address",
        "Status"
            });
            filterColumn.SelectedIndex = 0;

            // Create and configure the search label
            Label searchLabel = new Label
            {
                Text = "Search in:",
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(5)
            };

            // Add controls to search panel
            ((TableLayoutPanel)searchAndFilterPanel).Controls.Add(searchBox, 0, 0);
            ((TableLayoutPanel)searchAndFilterPanel).Controls.Add(searchLabel, 1, 0);
            ((TableLayoutPanel)searchAndFilterPanel).Controls.Add(filterColumn, 2, 0);

            // Modify the main layout to include the search panel
            // First, store the index of the grid panel
            int gridIndex = mainLayout.Controls.IndexOf(mainLayout.Controls.OfType<Panel>()
                .First(p => p.Controls.Contains(gridView)));

            // Add the new row and insert the search panel
            mainLayout.RowCount++;
            mainLayout.RowStyles.Insert(2, new RowStyle(SizeType.Absolute, 40));

            // Move all controls down one row starting from the grid's position
            for (int i = mainLayout.Controls.Count - 1; i >= 0; i--)
            {
                Control control = mainLayout.Controls[i];
                int row = mainLayout.GetRow(control);
                if (row >= 2)
                {
                    mainLayout.SetRow(control, row + 1);
                }
            }

            // Add the search panel
            mainLayout.Controls.Add(searchAndFilterPanel, 0, 2);

            // Add event handlers
            searchBox.TextChanged += (s, e) => FilterGrid();
            filterColumn.SelectedIndexChanged += (s, e) => FilterGrid();
        }

        private void FilterGrid()
        {
            string searchText = searchBox.Text.ToLower();
            string filterBy = filterColumn.SelectedItem.ToString();

            foreach (DataGridViewRow row in gridView.Rows)
            {
                bool visible = false;
                if (string.IsNullOrEmpty(searchText))
                {
                    visible = true;
                }
                else if (filterBy == "All Columns")
                {
                    visible = row.Cells.Cast<DataGridViewCell>()
                        .Any(cell => cell.Value?.ToString().ToLower()
                        .Contains(searchText) == true);
                }
                else
                {
                    int columnIndex = gridView.Columns.Cast<DataGridViewColumn>()
                        .FirstOrDefault(col => col.HeaderText == filterBy)?.Index ?? -1;

                    if (columnIndex >= 0)
                    {
                        visible = row.Cells[columnIndex].Value?.ToString()
                            .ToLower().Contains(searchText) == true;
                    }
                }
                row.Visible = visible;
            }
        }

        // Add these methods for the export functionality
        private async void ExportButton_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files|*.xlsx|CSV Files|*.csv";
                sfd.FileName = $"FIDS_Clients_{DateTime.Now:yyyyMMdd_HHmmss}";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        statusLabel.Text = "Exporting data...";
                        statusLabel.ForeColor = Color.Blue;
                        exportButton.Enabled = false;

                        if (sfd.FileName.EndsWith(".xlsx"))
                        {
                            await Task.Run(() => ExportToExcel(sfd.FileName));
                        }
                        else
                        {
                            await Task.Run(() => ExportToCsv(sfd.FileName));
                        }

                        MessageBox.Show("Export completed successfully!", "Export",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Export failed: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        exportButton.Enabled = true;
                        statusLabel.Text = "Export completed";
                        statusLabel.ForeColor = Color.Green;
                    }
                }
            }
        }

        private void ExportToExcel(string filePath)
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("FIDS Clients");

                // Add headers
                for (int i = 0; i < gridView.Columns.Count; i++)
                {
                    worksheet.Cell(1, i + 1).Value = gridView.Columns[i].HeaderText;
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }

                // Add data
                for (int i = 0; i < gridView.Rows.Count; i++)
                {
                    var row = gridView.Rows[i];
                    for (int j = 0; j < gridView.Columns.Count; j++)
                    {
                        worksheet.Cell(i + 2, j + 1).Value = row.Cells[j].Value?.ToString() ?? "";
                    }

                    // Color coding based on status
                    string status = row.Cells[4].Value?.ToString() ?? "";
                    var xlRow = worksheet.Row(i + 2);

                    switch (status)
                    {
                        case "Synchronized":
                            xlRow.Style.Fill.BackgroundColor = XLColor.LightGreen;
                            break;
                        case "Unreachable":
                        case "Connection Failed":
                            xlRow.Style.Fill.BackgroundColor = XLColor.LightPink;
                            break;
                        case "Connecting...":
                        case "Setting time...":
                            xlRow.Style.Fill.BackgroundColor = XLColor.LightYellow;
                            break;
                    }
                }

                // Auto-fit columns
                worksheet.Columns().AdjustToContents();

                // Add summary
                var summarySheet = workbook.Worksheets.Add("Summary");
                int totalClients = gridView.Rows.Count;
                int syncedClients = gridView.Rows.Cast<DataGridViewRow>()
                    .Count(r => r.Cells[4].Value?.ToString() == "Synchronized");
                int failedClients = gridView.Rows.Cast<DataGridViewRow>()
                    .Count(r => r.Cells[4].Value?.ToString().Contains("Failed") == true);

                summarySheet.Cell("A1").Value = "Export Summary";
                summarySheet.Cell("A2").Value = "Total Clients:";
                summarySheet.Cell("B2").Value = totalClients;
                summarySheet.Cell("A3").Value = "Synchronized Clients:";
                summarySheet.Cell("B3").Value = syncedClients;
                summarySheet.Cell("A4").Value = "Failed Clients:";
                summarySheet.Cell("B4").Value = failedClients;
                summarySheet.Cell("A5").Value = "Export Date:";
                summarySheet.Cell("B5").Value = DateTime.Now.ToString();

                workbook.SaveAs(filePath);
            }
        }

        private void ExportToCsv(string filePath)
        {
            using (var writer = new StreamWriter(filePath))
            {
                // Write headers
                writer.WriteLine(string.Join(",", gridView.Columns.Cast<DataGridViewColumn>()
                    .Select(column => $"\"{column.HeaderText}\"")));

                // Write rows
                foreach (DataGridViewRow row in gridView.Rows)
                {
                    writer.WriteLine(string.Join(",", row.Cells.Cast<DataGridViewCell>()
                        .Select(cell => $"\"{cell.Value?.ToString() ?? ""}\"")));
                }
            }
        }

        // Add these new methods for statistics functionality
        private void StatisticsButton_Click(object sender, EventArgs e)
        {
            var stats = CalculateStatistics();
            ShowStatisticsDialog(stats);
        }

        private ClientStatistics CalculateStatistics()
        {
            var stats = new ClientStatistics
            {
                TotalClients = clients.Count,
                ClientsBySystem = new Dictionary<string, int>(),
                ClientsByStatus = new Dictionary<string, int>(),
                LastRefresh = DateTime.Now,
                DriftStats = new Dictionary<string, int>
        {
            { "Normal", 0 },
            { "Critical", 0 },
            { "Unknown", 0 }
        },
                CriticalDriftClients = new List<(string Name, TimeSpan Drift)>()
            };

            // Calculate drift statistics
            foreach (var client in clients)
            {
                if (!client.LastCheckTime.HasValue)
                {
                    stats.DriftStats["Unknown"]++;
                }
                else if (client.IsDriftCritical)
                {
                    stats.DriftStats["Critical"]++;
                    if (client.TimeDrift.HasValue)
                    {
                        stats.CriticalDriftClients.Add((client.Name, client.TimeDrift.Value));
                    }
                }
                else
                {
                    stats.DriftStats["Normal"]++;
                }
            }

            // System and status statistics from grid
            foreach (DataGridViewRow row in gridView.Rows)
            {
                var status = row.Cells[4].Value?.ToString() ?? "Unknown";
                var macExpected = row.Cells[2].Value?.ToString();
                var macActual = row.Cells[3].Value?.ToString();
                var clientName = row.Cells[0].Value?.ToString() ?? "Unknown";
                var monitor = row.Cells[7].Value?.ToString() ?? "Unknown";
                var type = row.Cells[8].Value?.ToString() ?? "Unknown";
                var group = row.Cells[9].Value?.ToString() ?? "Unknown";

                // Update system counts (using Type as system category)
                if (!stats.ClientsBySystem.ContainsKey(type))
                    stats.ClientsBySystem[type] = 0;
                stats.ClientsBySystem[type]++;

                // Update status counts
                if (!stats.ClientsByStatus.ContainsKey(status))
                    stats.ClientsByStatus[status] = 0;
                stats.ClientsByStatus[status]++;

                // Update online/offline counts
                switch (status.ToLower())
                {
                    case "connected":
                    case "synchronized":
                        stats.OnlineClients++;
                        break;
                    case "unreachable":
                    case "connection failed":
                        stats.OfflineClients++;
                        break;
                }

                if (status.Equals("Synchronized", StringComparison.OrdinalIgnoreCase))
                    stats.SynchronizedClients++;

                // Update MAC match counts
                if (!string.IsNullOrEmpty(macExpected) && !string.IsNullOrEmpty(macActual) &&
                    macExpected != "-" && macActual != "-")
                {
                    if (macExpected.Equals(macActual, StringComparison.OrdinalIgnoreCase))
                        stats.MatchedMacs++;
                    else
                        stats.MismatchedMacs++;
                }
            }

            return stats;
        }

        private void ShowStatisticsDialog(ClientStatistics stats)
        {
            var statsForm = new Form
            {
                Text = "FIDS Client Statistics",
                Size = new Size(800, 600),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };

            var tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Location = new Point(10, 10),  // Fixed: Using Point instead of Padding
                Size = new Size(statsForm.ClientSize.Width - 20, statsForm.ClientSize.Height - 50)
            };

            // Summary Tab
            var summaryTab = new TabPage("Summary");
            var summaryText = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10),
                ScrollBars = ScrollBars.Both
            };

            // Add summary content
            summaryText.AppendText($"FIDS Client Statistics Summary\r\n");
            summaryText.AppendText($"Generated: {stats.LastRefresh:yyyy-MM-dd HH:mm:ss}\r\n\r\n");
            summaryText.AppendText($"Total Clients: {stats.TotalClients}\r\n");
            summaryText.AppendText($"Online Clients: {stats.OnlineClients}\r\n");
            summaryText.AppendText($"Offline Clients: {stats.OfflineClients}\r\n");
            summaryText.AppendText($"Synchronized Clients: {stats.SynchronizedClients}\r\n");
            summaryText.AppendText($"Connection Success Rate: {(stats.OnlineClients * 100.0 / Math.Max(stats.TotalClients, 1)):F1}%\r\n\r\n");

            summaryText.AppendText("MAC Address Statistics:\r\n");
            summaryText.AppendText($"Matched MACs: {stats.MatchedMacs}\r\n");
            summaryText.AppendText($"Mismatched MACs: {stats.MismatchedMacs}\r\n");

            // Details Tab
            var detailsTab = new TabPage("System Details");
            var detailsGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            detailsGrid.Columns.Add("System", "System");
            detailsGrid.Columns.Add("Clients", "Clients");
            detailsGrid.Columns.Add("Percentage", "Percentage");

            foreach (var system in stats.ClientsBySystem.OrderByDescending(x => x.Value))
            {
                double percentage = (system.Value * 100.0) / Math.Max(stats.TotalClients, 1);
                detailsGrid.Rows.Add(system.Key, system.Value, $"{percentage:F1}%");
            }

            // Status Tab
            var statusTab = new TabPage("Status Distribution");
            var statusGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            statusGrid.Columns.Add("Status", "Status");
            statusGrid.Columns.Add("Count", "Count");
            statusGrid.Columns.Add("Percentage", "Percentage");

            foreach (var status in stats.ClientsByStatus.OrderByDescending(x => x.Value))
            {
                double percentage = (status.Value * 100.0) / Math.Max(stats.TotalClients, 1);
                statusGrid.Rows.Add(status.Key, status.Value, $"{percentage:F1}%");
            }

            // Add controls to tabs
            summaryTab.Controls.Add(summaryText);
            detailsTab.Controls.Add(detailsGrid);
            statusTab.Controls.Add(statusGrid);

            // Add tabs to control
            tabControl.TabPages.Add(summaryTab);
            tabControl.TabPages.Add(detailsTab);
            tabControl.TabPages.Add(statusTab);

            // Create refresh button
            var refreshButton = new Button
            {
                Text = "Refresh",
                Dock = DockStyle.Bottom,
                Height = 30
            };
            refreshButton.Click += (s, e) =>
            {
                var newStats = CalculateStatistics();
                ShowStatisticsDialog(newStats);
                statsForm.Close();
            };

            // Add drift statistics to summary text
            summaryText.AppendText("\r\nTime Drift Statistics:\r\n");
            summaryText.AppendText($"Normal Drift: {stats.DriftStats["Normal"]}\r\n");
            summaryText.AppendText($"Critical Drift: {stats.DriftStats["Critical"]}\r\n");
            summaryText.AppendText($"Unknown: {stats.DriftStats["Unknown"]}\r\n");

            if (stats.CriticalDriftClients.Any())
            {
                summaryText.AppendText("\r\nClients with Critical Drift:\r\n");
                foreach (var (name, drift) in stats.CriticalDriftClients)
                {
                    summaryText.AppendText($"- {name}: {drift.TotalSeconds:F2} seconds\r\n");
                }
            }

            // Add panel to contain tabControl and maintain padding
            var containerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            containerPanel.Controls.Add(tabControl);

            // Add controls to form
            statsForm.Controls.Add(containerPanel);
            statsForm.Controls.Add(refreshButton);

            // Show the form
            statsForm.ShowDialog();
        }

        private void GridView_SelectionChanged(object sender, EventArgs e)
        {
            batchOperationsButton.Enabled = gridView.SelectedRows.Count > 0;
            modifyButton.Enabled = gridView.SelectedRows.Count == 1;
        }

        private void BatchOperationsButton_Click(object sender, EventArgs e)
        {
            var selectedRows = gridView.SelectedRows.Cast<DataGridViewRow>().ToList();
            ShowBatchOperationsDialog(selectedRows);
        }

        private void ShowBatchOperationsDialog(List<DataGridViewRow> selectedRows)
        {
            var batchForm = new Form
            {
                Text = "Batch Operations",
                Size = new Size(600, 500),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };

            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10)
            };

            // Selected clients list
            var clientsList = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.None
            };

            foreach (var row in selectedRows)
            {
                clientsList.Items.Add($"{row.Cells[0].Value} ({row.Cells[1].Value})");
            }

            // Command input
            var commandPanel = new GroupBox
            {
                Text = "Command",
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            var commandBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Both
            };

            commandPanel.Controls.Add(commandBox);

            // Progress panel
            var progressPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 60
            };

            var progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Continuous,
                Height = 25,
                Width = 550,
                Location = new Point(10, 5)
            };

            var progressLabel = new Label
            {
                AutoSize = true,
                Location = new Point(10, 35)
            };

            progressPanel.Controls.AddRange(new Control[] { progressBar, progressLabel });

            // Buttons panel
            var buttonsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 40
            };

            var executeButton = new Button
            {
                Text = "Execute",
                Width = 100,
                Height = 30
            };

            var syncTimeButton = new Button
            {
                Text = "Sync Time",
                Width = 100,
                Height = 30
            };

            var closeButton = new Button
            {
                Text = "Close",
                Width = 100,
                Height = 30
            };

            buttonsPanel.Controls.AddRange(new Control[] { closeButton, syncTimeButton, executeButton });

            // Add all panels to main layout
            mainPanel.Controls.Add(new Label { Text = $"Selected Clients ({selectedRows.Count}):", Height = 20 }, 0, 0);
            mainPanel.Controls.Add(clientsList, 0, 1);
            mainPanel.Controls.Add(commandPanel, 0, 2);
            mainPanel.Controls.Add(progressPanel, 0, 3);
            mainPanel.Controls.Add(buttonsPanel, 0, 4);

            // Configure row styles
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 25));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            // Add event handlers
            executeButton.Click += async (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(commandBox.Text))
                {
                    MessageBox.Show("Please enter a command.", "Batch Operations",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                await ExecuteBatchCommand(selectedRows, commandBox.Text, progressBar, progressLabel);
            };

            syncTimeButton.Click += async (s, e) =>
            {
                await SynchronizeSelectedClients(selectedRows, progressBar, progressLabel);
            };

            closeButton.Click += (s, e) => batchForm.Close();

            batchForm.Controls.Add(mainPanel);
            batchForm.ShowDialog();
        }

        private async Task ExecuteBatchCommand(List<DataGridViewRow> selectedRows, string command,
            ProgressBar progressBar, Label progressLabel)
        {
            progressBar.Maximum = selectedRows.Count;
            progressBar.Value = 0;
            int successCount = 0;
            int failCount = 0;

            foreach (var row in selectedRows)
            {
                string ip = row.Cells[1].Value.ToString();
                string name = row.Cells[0].Value.ToString();
                progressLabel.Text = $"Processing: {name}";

                LogConnection(name, ip, "Batch Command", "Starting", command);
                try
                {
                    using (var sshClient = new SshClient(ip, "root", "123456"))
                    {
                        await Task.Run(() => sshClient.Connect());
                        var result = await Task.Run(() => sshClient.RunCommand(command));

                        if (result.ExitStatus == 0)
                            successCount++;
                        else
                            failCount++;

                        sshClient.Disconnect();
                        LogConnection(name, ip, "Batch Command", "Completed", result.Result);
                    }

                }
                catch (Exception ex)
                {
                    LogConnection(name, ip, "Batch Command", "Failed", ex.Message);
                    failCount++;
                    LogError($"Batch command failed for {name}: {ex.Message}");
                }

                progressBar.Value++;
                progressLabel.Text = $"Completed: {progressBar.Value}/{selectedRows.Count} " +
                                   $"(Success: {successCount}, Failed: {failCount})";
            }

            progressLabel.Text = $"Operation completed. Success: {successCount}, Failed: {failCount}";
            MessageBox.Show($"Batch operation completed.\nSuccess: {successCount}\nFailed: {failCount}",
                "Batch Operation Result", MessageBoxButtons.OK,
                failCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        private async Task SynchronizeSelectedClients(List<DataGridViewRow> selectedRows,
            ProgressBar progressBar, Label progressLabel)
        {
            progressBar.Maximum = selectedRows.Count;
            progressBar.Value = 0;
            int successCount = 0;
            int failCount = 0;
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            foreach (var row in selectedRows)
            {
                string ip = row.Cells[1].Value.ToString();
                string name = row.Cells[0].Value.ToString();
                progressLabel.Text = $"Syncing: {name}";

                try
                {
                    using (var sshClient = new SshClient(ip, "root", "123456"))
                    {
                        await Task.Run(() => sshClient.Connect());

                        string dateCommand = $"date -s \"{currentTime}\"";
                        await Task.Run(() => sshClient.RunCommand(dateCommand));
                        await Task.Run(() => sshClient.RunCommand("hwclock --systohc"));

                        successCount++;
                        UpdateGridRow(name, ip, "Synchronized", currentTime);
                        sshClient.Disconnect();
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    UpdateGridRow(name, ip, "Sync Failed");
                    LogError($"Batch sync failed for {name}: {ex.Message}");
                }

                progressBar.Value++;
                progressLabel.Text = $"Completed: {progressBar.Value}/{selectedRows.Count} " +
                                   $"(Success: {successCount}, Failed: {failCount})";
            }

            progressLabel.Text = $"Synchronization completed. Success: {successCount}, Failed: {failCount}";
            MessageBox.Show($"Batch synchronization completed.\nSuccess: {successCount}\nFailed: {failCount}",
                "Batch Sync Result", MessageBoxButtons.OK,
                failCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        private void LogError(string message)
        {
            // You can implement more sophisticated logging here
            Debug.WriteLine($"[ERROR] {DateTime.Now}: {message}");
        }

        private void LogConnection(string clientName, string ip, string action, string status, string details = "")
        {
            var log = new ConnectionLog
            {
                Timestamp = DateTime.Now,
                ClientName = clientName,
                IP = ip,
                Action = action,
                Status = status,
                Details = details
            };

            connectionHistory.Add(log);
        }

        private void HistoryButton_Click(object sender, EventArgs e)
        {
            ShowConnectionHistory();
        }

        private void ShowConnectionHistory()
        {
            var historyForm = new Form
            {
                Text = "Connection History",
                Size = new Size(1000, 600),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = true,
                FormBorderStyle = FormBorderStyle.Sizable
            };

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };

            // Filter panel
            var filterPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Height = 30
            };

            var clientFilter = new ComboBox
            {
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDown,
                Text = "Filter by Client..."
            };

            var actionFilter = new ComboBox
            {
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDown,
                Text = "Filter by Action..."
            };

            var statusFilter = new ComboBox
            {
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDown,
                Text = "Filter by Status..."
            };

            var clearFiltersButton = new Button
            {
                Text = "Clear Filters",
                Width = 100
            };

            filterPanel.Controls.AddRange(new Control[] {
        new Label { Text = "Client:", Padding = new Padding(5, 5, 5, 0) },
        clientFilter,
        new Label { Text = "Action:", Padding = new Padding(5, 5, 5, 0) },
        actionFilter,
        new Label { Text = "Status:", Padding = new Padding(5, 5, 5, 0) },
        statusFilter,
        clearFiltersButton
    });

            // Grid
            var historyGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = true
            };

            historyGrid.Columns.AddRange(new DataGridViewColumn[]
            {
        new DataGridViewTextBoxColumn { Name = "Timestamp", HeaderText = "Timestamp", FillWeight = 15 },
        new DataGridViewTextBoxColumn { Name = "ClientName", HeaderText = "Client Name", FillWeight = 20 },
        new DataGridViewTextBoxColumn { Name = "IP", HeaderText = "IP Address", FillWeight = 15 },
        new DataGridViewTextBoxColumn { Name = "Action", HeaderText = "Action", FillWeight = 15 },
        new DataGridViewTextBoxColumn { Name = "Status", HeaderText = "Status", FillWeight = 15 },
        new DataGridViewTextBoxColumn { Name = "Details", HeaderText = "Details", FillWeight = 20 }
            });

            // Button panel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 40
            };

            var exportButton = new Button { Text = "Export History", Width = 120 };
            var refreshButton = new Button { Text = "Refresh", Width = 100 };
            var clearButton = new Button { Text = "Clear History", Width = 120 };

            buttonPanel.Controls.AddRange(new Control[] { clearButton, exportButton, refreshButton });

            // Add controls to main layout
            mainLayout.Controls.Add(filterPanel, 0, 0);
            mainLayout.Controls.Add(historyGrid, 0, 1);
            mainLayout.Controls.Add(buttonPanel, 0, 2);

            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));

            // Populate data
            void RefreshGrid(string clientFilter = "", string actionFilter = "", string statusFilter = "")
            {
                historyGrid.Rows.Clear();
                var filteredHistory = connectionHistory
                    .Where(log =>
                        (string.IsNullOrEmpty(clientFilter) || log.ClientName.Contains(clientFilter)) &&
                        (string.IsNullOrEmpty(actionFilter) || log.Action.Contains(actionFilter)) &&
                        (string.IsNullOrEmpty(statusFilter) || log.Status.Contains(statusFilter)))
                    .OrderByDescending(log => log.Timestamp)
                    .ToList();

                foreach (var log in filteredHistory)
                {
                    historyGrid.Rows.Add(
                        log.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                        log.ClientName,
                        log.IP,
                        log.Action,
                        log.Status,
                        log.Details
                    );
                }

                // Update filter dropdowns
                UpdateFilterDropdowns();
            }

            void UpdateFilterDropdowns()
            {
                // Update client filter
                var clients = connectionHistory.Select(l => l.ClientName).Distinct().OrderBy(c => c).ToList();
                clientFilter.Items.Clear();
                clientFilter.Items.AddRange(clients.Cast<object>().ToArray());

                // Update action filter
                var actions = connectionHistory.Select(l => l.Action).Distinct().OrderBy(a => a).ToList();
                actionFilter.Items.Clear();
                actionFilter.Items.AddRange(actions.Cast<object>().ToArray());

                // Update status filter
                var statuses = connectionHistory.Select(l => l.Status).Distinct().OrderBy(s => s).ToList();
                statusFilter.Items.Clear();
                statusFilter.Items.AddRange(statuses.Cast<object>().ToArray());
            }

            // Add event handlers
            clientFilter.TextChanged += (s, e) => RefreshGrid(clientFilter.Text, actionFilter.Text, statusFilter.Text);
            actionFilter.TextChanged += (s, e) => RefreshGrid(clientFilter.Text, actionFilter.Text, statusFilter.Text);
            statusFilter.TextChanged += (s, e) => RefreshGrid(clientFilter.Text, actionFilter.Text, statusFilter.Text);
            clearFiltersButton.Click += (s, e) =>
            {
                clientFilter.Text = "";
                actionFilter.Text = "";
                statusFilter.Text = "";
                RefreshGrid();
            };

            exportButton.Click += async (s, e) =>
            {
                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Filter = "Excel Files|*.xlsx|CSV Files|*.csv";
                    sfd.FileName = $"ConnectionHistory_{DateTime.Now:yyyyMMdd_HHmmss}";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        await ExportHistory(sfd.FileName, historyGrid.Rows.Cast<DataGridViewRow>()
                            .Where(r => r.Visible).ToList());
                    }
                }
            };

            refreshButton.Click += (s, e) => RefreshGrid(clientFilter.Text, actionFilter.Text, statusFilter.Text);

            clearButton.Click += (s, e) =>
            {
                if (MessageBox.Show("Are you sure you want to clear all history?", "Clear History",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    connectionHistory.Clear();
                    RefreshGrid();
                }
            };

            // Initial load
            RefreshGrid();

            historyForm.Controls.Add(mainLayout);
            historyForm.Show();
        }

        private async Task ExportHistory(string filePath, List<DataGridViewRow> rows)
        {
            try
            {
                if (filePath.EndsWith(".xlsx"))
                {
                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var worksheet = workbook.Worksheets.Add("Connection History");

                        // Add headers
                        var headers = new[] { "Timestamp", "Client Name", "IP Address", "Action", "Status", "Details" };
                        for (int i = 0; i < headers.Length; i++)
                        {
                            worksheet.Cell(1, i + 1).Value = headers[i];
                        }

                        // Add data
                        for (int i = 0; i < rows.Count; i++)
                        {
                            var row = rows[i];
                            for (int j = 0; j < row.Cells.Count; j++)
                            {
                                worksheet.Cell(i + 2, j + 1).Value = row.Cells[j].Value?.ToString() ?? "";
                            }
                        }

                        worksheet.Columns().AdjustToContents();
                        await Task.Run(() => workbook.SaveAs(filePath));
                    }
                }
                else
                {
                    using (var writer = new StreamWriter(filePath))
                    {
                        // Write headers
                        writer.WriteLine(string.Join(",", new[] { "Timestamp", "Client Name", "IP Address", "Action", "Status", "Details" }
                            .Select(h => $"\"{h}\"")));

                        // Write data
                        foreach (var row in rows)
                        {
                            writer.WriteLine(string.Join(",", row.Cells.Cast<DataGridViewCell>()
                                .Select(cell => $"\"{cell.Value?.ToString() ?? ""}\"")));
                        }
                    }
                }

                MessageBox.Show("Export completed successfully!", "Export",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task CheckTimeDrift(ClientInfo client)
        {
            try
            {
                using (var sshClient = new SshClient(client.IP, "root", "123456"))
                {
                    await Task.Run(() => sshClient.Connect());

                    // Get client's current time
                    var timeResult = await Task.Run(() =>
                        sshClient.RunCommand("date '+%Y-%m-%d %H:%M:%S'").Result.Trim());

                    if (DateTime.TryParse(timeResult, out DateTime clientTime))
                    {
                        var serverTime = DateTime.Now;
                        var drift = clientTime - serverTime;

                        // Update client info
                        client.LastCheckTime = serverTime;
                        client.TimeDrift = drift;
                        client.IsDriftCritical = Math.Abs(drift.TotalSeconds) > 10;

                        // Add to history with formatted display
                        string driftText = $"{(drift.TotalSeconds >= 0 ? "+" : "")}{drift.TotalSeconds:F2} sec";
                        client.DriftHistory.Add(new TimeDriftRecord
                        {
                            CheckTime = serverTime,
                            Drift = drift,
                            Status = client.IsDriftCritical ? "Critical" : "Normal",
                            DisplayTime = serverTime.ToString("HH:mm:ss")
                        });

                        // Update grid with drift information
                        UpdateGridRowWithDrift(client.Name, client.IP, drift);

                        // Log drift check with formatted display
                        LogConnection(client.Name, client.IP, "Drift Check",
                            client.IsDriftCritical ? "Critical Drift" : "Normal",
                            $"Time: {serverTime:HH:mm:ss} || Drift: {driftText}");
                    }

                    sshClient.Disconnect();
                }
            }
            catch (Exception ex)
            {
                LogError($"Drift check failed for {client.Name}: {ex.Message}");
            }
        }

        private void UpdateGridRowWithDrift(string name, string ip, TimeSpan drift)
        {
            if (gridView.InvokeRequired)
            {
                gridView.Invoke(new Action(() => UpdateGridRowWithDrift(name, ip, drift)));
                return;
            }

            foreach (DataGridViewRow row in gridView.Rows)
            {
                if (row.Cells[0].Value.ToString() == name && row.Cells[1].Value.ToString() == ip)
                {
                    // Format the time and drift
                    string currentTime = DateTime.Now.ToString("HH:mm:ss");
                    string driftText = $"{(drift.TotalSeconds >= 0 ? "+" : "")}{drift.TotalSeconds:F2} sec";
                    string displayText = $"{currentTime}  ||  {driftText}";

                    // Update current time column with formatted time and drift
                    if (Math.Abs(drift.TotalSeconds) > 10)
                    {
                        row.Cells[5].Style.BackColor = Color.LightPink;
                        row.Cells[5].Style.ForeColor = Color.Red;
                        displayText += " ⚠"; // Add warning symbol for critical drift
                    }
                    else
                    {
                        row.Cells[5].Style.BackColor = Color.LightGreen;
                        row.Cells[5].Style.ForeColor = Color.Black;
                    }
                    row.Cells[5].Value = displayText;
                    break;
                }
            }
        }

        // Add this method for viewing drift history
        private void ShowDriftHistory(ClientInfo client)
        {
            var historyForm = new Form
            {
                Text = $"Time Drift History - {client.Name}",
                Size = new Size(600, 400),
                StartPosition = FormStartPosition.CenterParent
            };

            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true
            };

            grid.Columns.Add("Time", "Time");
            grid.Columns.Add("Drift", "Drift");
            grid.Columns.Add("Status", "Status");

            foreach (var record in client.DriftHistory.OrderByDescending(h => h.CheckTime))
            {
                string driftText = $"{(record.Drift.TotalSeconds >= 0 ? "+" : "")}{record.Drift.TotalSeconds:F2} sec";

                var row = grid.Rows.Add(
                    record.DisplayTime,
                    driftText,
                    record.Status
                );

                if (record.Status == "Critical")
                {
                    grid.Rows[row].DefaultCellStyle.BackColor = Color.LightPink;
                    grid.Rows[row].DefaultCellStyle.ForeColor = Color.Red;
                }
            }

            historyForm.Controls.Add(grid);
            historyForm.ShowDialog();
        }

        private void AddDriftHistoryContextMenu()
        {
            var contextMenu = new ContextMenuStrip();
            var viewDriftHistoryItem = new ToolStripMenuItem("View Drift History");

            viewDriftHistoryItem.Click += (s, e) =>
            {
                if (gridView.SelectedRows.Count > 0)
                {
                    var row = gridView.SelectedRows[0];
                    var client = clients.FirstOrDefault(c =>
                        c.Name == row.Cells[0].Value.ToString() &&
                        c.IP == row.Cells[1].Value.ToString());

                    if (client != null)
                    {
                        ShowDriftHistory(client);
                    }
                }
            };

            contextMenu.Items.Add(viewDriftHistoryItem);
            gridView.ContextMenuStrip = contextMenu;
        }

        // Add this to your Form1 constructor or InitializeUI method
        private void InitializeTimeDriftTracking()
        {
            AddDriftHistoryContextMenu();
        }

        private void AddFilterControls()
        {
            try
            {
                mainLayout.SuspendLayout();

                // Increment RowCount and adjust RowStyles
                mainLayout.RowCount++;

                // Insert new RowStyle for filter panel
                var currentStyles = mainLayout.RowStyles.Cast<RowStyle>().ToList();
                mainLayout.RowStyles.Clear();
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // Status label
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // Filter panel
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // Grid
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));  // Buttons

                // Create filter panel
                var filterPanel = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    Height = 40,
                    ColumnCount = 11,
                    RowCount = 1,
                    Padding = new Padding(5)
                };

                // Set equal column widths
                for (int i = 0; i < 11; i++)
                {
                    filterPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 9.09f));
                }

                // Initialize filters
                monitorFilter = CreateFilterComboBox("Monitor");
                typeFilter = CreateFilterComboBox("Type");
                groupFilter = CreateFilterComboBox("Group");
                levelFilter = CreateFilterComboBox("Level");
                zoneFilter = CreateFilterComboBox("Zone");

                // Add filters to panel
                int column = 0;
                AddFilterToPanel(filterPanel, "Monitor:", monitorFilter, ref column);
                AddFilterToPanel(filterPanel, "Type:", typeFilter, ref column);
                AddFilterToPanel(filterPanel, "Group:", groupFilter, ref column);
                AddFilterToPanel(filterPanel, "Level:", levelFilter, ref column);
                AddFilterToPanel(filterPanel, "Zone:", zoneFilter, ref column);

                // Add clear filters button
                var clearButton = new Button
                {
                    Text = "Clear Filters",
                    Width = 100,
                    Anchor = AnchorStyles.None
                };
                clearButton.Click += (s, e) => ClearFilters();
                filterPanel.Controls.Add(clearButton, column, 0);

                // Move existing controls down
                var controlsToMove = mainLayout.Controls.Cast<Control>()
                    .Where(c => mainLayout.GetRow(c) >= 1)
                    .OrderByDescending(c => mainLayout.GetRow(c))
                    .ToList();

                foreach (var control in controlsToMove)
                {
                    mainLayout.SetRow(control, mainLayout.GetRow(control) + 1);
                }

                // Add filter panel
                mainLayout.Controls.Add(filterPanel, 0, 1);

                mainLayout.ResumeLayout(true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in AddFilterControls: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show("Error initializing filters. Check the application log for details.",
                    "Initialization Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private ComboBox CreateFilterComboBox(string name)
        {
            var comboBox = new ComboBox
            {
                Name = $"{name.ToLower()}Filter",
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 120,
                Dock = DockStyle.Fill,
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };

            comboBox.Items.Add($"All {name}s");
            comboBox.SelectedIndex = 0;
            comboBox.SelectedIndexChanged += (s, e) => ApplyFilters();

            return comboBox;
        }

        private void AddFilterToPanel(TableLayoutPanel panel, string labelText, ComboBox comboBox, ref int column)
        {
            var label = new Label
            {
                Text = labelText,
                TextAlign = ContentAlignment.MiddleRight,
                Anchor = AnchorStyles.Right,
                AutoSize = true
            };

            panel.Controls.Add(label, column++, 0);
            panel.Controls.Add(comboBox, column++, 0);
        }

        void UpdateFilter(ComboBox filter, Func<ClientInfo, string> selector)
        {
            if (filter == null) throw new ArgumentNullException(nameof(filter));

            var currentSelection = filter.SelectedItem?.ToString();
            var items = clients
                .Select(selector)
                .Where(item => !string.IsNullOrEmpty(item))
                .Distinct()
                .OrderBy(item => item)
                .ToList();

            filter.Items.Clear();
            filter.Items.Add($"All {filter.Name.Replace("Filter", "")}s");
            filter.Items.AddRange(items.Cast<object>().ToArray());

            if (currentSelection != null && filter.Items.Contains(currentSelection))
                filter.SelectedItem = currentSelection;
            else
                filter.SelectedIndex = 0;
        }

        private void UpdateComboBoxItems(ComboBox comboBox, IEnumerable<string> items)
        {
            var currentSelection = comboBox.SelectedItem?.ToString();
            comboBox.Items.Clear();
            comboBox.Items.Add($"All {comboBox.Name.Replace("Filter", "")}s");
            comboBox.Items.AddRange(items.Where(i => !string.IsNullOrEmpty(i)).OrderBy(i => i).ToArray());
            comboBox.SelectedItem = currentSelection ?? comboBox.Items[0];
        }

        private void ApplyFilters()
        {
            if (gridView.Rows.Count == 0) return;

            foreach (DataGridViewRow row in gridView.Rows)
            {
                bool visible = true;

                // Helper function to check filter condition
                bool CheckFilter(ComboBox filter, string columnName)
                {
                    return filter.SelectedIndex == 0 ||
                           row.Cells[columnName].Value?.ToString() == filter.SelectedItem.ToString();
                }

                visible &= CheckFilter(monitorFilter, "Monitor");
                visible &= CheckFilter(typeFilter, "Type");
                visible &= CheckFilter(groupFilter, "Group");
                visible &= CheckFilter(levelFilter, "Level");
                visible &= CheckFilter(zoneFilter, "Zone");

                row.Visible = visible;
            }

            // Update any status or count displays
            UpdateFilterStatus();
        }

        private void UpdateFilterStatus()
        {
            int visibleRows = gridView.Rows.Cast<DataGridViewRow>().Count(r => r.Visible);
            statusLabel.Text = $"Showing {visibleRows} of {gridView.Rows.Count} clients";
        }

        private void ClearFilters()
        {
            monitorFilter.SelectedIndex = 0;
            typeFilter.SelectedIndex = 0;
            groupFilter.SelectedIndex = 0;
            levelFilter.SelectedIndex = 0;
            zoneFilter.SelectedIndex = 0;

            // No need to call ApplyFilters() explicitly as it will be triggered by the SelectedIndexChanged events
        }

        private async void LoadClientsFromExcel()
        {
            try
            {
                if (!File.Exists(EXCEL_FILE_PATH))
                {
                    MessageBox.Show("Excel file not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    statusLabel.Text = "Excel file not found";
                    statusLabel.ForeColor = Color.Red;
                    return;
                }

                statusLabel.Text = "Loading clients...";
                statusLabel.ForeColor = Color.Blue;

                await Task.Run(() => LoadIPsFromExcel(EXCEL_FILE_PATH));

                if (clients.Any())
                {
                    syncButton.Enabled = true;
                    statusLabel.Text = $"Loaded {clients.Count} clients";
                    statusLabel.ForeColor = Color.Green;
                    LogLoadingSummary();
                }
                else
                {
                    statusLabel.Text = "No clients found in file";
                    statusLabel.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Error loading clients";
                statusLabel.ForeColor = Color.Red;
                Debug.WriteLine($"Excel loading error: {ex}");
            }
        }

        private void LogLoadingSummary()
        {
            var summary = new StringBuilder();
            summary.AppendLine($"Total clients loaded: {clients.Count}");
            summary.AppendLine("\nDistinct values found:");

            summary.AppendLine($"Monitors: {clients.Select(c => c.Monitor).Where(m => !string.IsNullOrEmpty(m)).Distinct().Count()}");
            summary.AppendLine($"Types: {clients.Select(c => c.Type).Where(t => !string.IsNullOrEmpty(t)).Distinct().Count()}");
            summary.AppendLine($"Groups: {clients.Select(c => c.Group).Where(g => !string.IsNullOrEmpty(g)).Distinct().Count()}");
            summary.AppendLine($"Levels: {clients.Select(c => c.Level).Where(l => !string.IsNullOrEmpty(l)).Distinct().Count()}");
            summary.AppendLine($"Zones: {clients.Select(c => c.Zone).Where(z => !string.IsNullOrEmpty(z)).Distinct().Count()}");

            Debug.WriteLine("Loading Summary:");
            Debug.WriteLine(summary.ToString());
        }
        private void UpdateFilterOptions()
        {
            if (clients == null) return;

            void UpdateFilter(ComboBox filter, Func<ClientInfo, string> selector)
            {
                var currentSelection = filter.SelectedItem?.ToString();
                var items = clients
                    .Select(selector)
                    .Where(item => !string.IsNullOrEmpty(item))
                    .Distinct()
                    .OrderBy(item => item)
                    .ToList();

                filter.Items.Clear();
                filter.Items.Add($"All {filter.Name.Replace("Filter", "")}s");
                filter.Items.AddRange(items.Cast<object>().ToArray());

                if (currentSelection != null && filter.Items.Contains(currentSelection))
                    filter.SelectedItem = currentSelection;
                else
                    filter.SelectedIndex = 0;
            }

            UpdateFilter(monitorFilter, c => c.Monitor);
            UpdateFilter(typeFilter, c => c.Type);
            UpdateFilter(groupFilter, c => c.Group);
            UpdateFilter(levelFilter, c => c.Level);
            UpdateFilter(zoneFilter, c => c.Zone);
        }

        // Add this method to create the buttons and add them to the interface
        private void AddCheckButtons()
        {
            // Create panel for check buttons
            var checkButtonsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Height = 40
            };

            // Configure columns to be equal width
            checkButtonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            checkButtonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            // Create buttons
            checkTimeButton = new Button
            {
                Text = "Check Current Time",
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 5, 10, 5)
            };
            checkTimeButton.Click += CheckTimeButton_Click;

            checkMacButton = new Button
            {
                Text = "Check Actual MAC",
                Dock = DockStyle.Fill,
                Margin = new Padding(10, 5, 10, 5)
            };
            checkMacButton.Click += CheckMacButton_Click;

            // Add buttons to panel
            checkButtonsPanel.Controls.Add(checkTimeButton, 0, 0);
            checkButtonsPanel.Controls.Add(checkMacButton, 1, 0);

            // Add the panel to mainLayout
            // Insert a new row for the check buttons after the search panel
            mainLayout.RowCount++;
            mainLayout.RowStyles.Insert(3, new RowStyle(SizeType.Absolute, 40));

            // Move all controls below the search panel down one row
            for (int i = mainLayout.Controls.Count - 1; i >= 0; i--)
            {
                Control control = mainLayout.Controls[i];
                int currentRow = mainLayout.GetRow(control);
                if (currentRow >= 3)
                {
                    mainLayout.SetRow(control, currentRow + 1);
                }
            }

            mainLayout.Controls.Add(checkButtonsPanel, 0, 3);
        }

        // Add these event handlers for the new buttons
        private void CheckTimeButton_Click(object sender, EventArgs e)
        {
            ShowCheckOptionsDialog("Time", async (clientsToCheck) =>
            {
                foreach (var client in clientsToCheck)
                {
                    try
                    {
                        var ping = new Ping();
                        var reply = await ping.SendPingAsync(client.IP, 1000);

                        if (reply.Status != IPStatus.Success)
                        {
                            UpdateGridRow(client.Name, client.IP, "Unreachable");
                            continue;
                        }

                        using (var sshClient = new SshClient(client.IP, "root", "123456"))
                        {
                            await Task.Run(() => sshClient.Connect());
                            await CheckTimeDrift(client);
                            sshClient.Disconnect();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogConnection(client.Name, client.IP, "Check Time", "Failed", ex.Message);
                        UpdateGridRow(client.Name, client.IP, "Check Time Failed");
                    }
                }
            });
        }

        private void CheckMacButton_Click(object sender, EventArgs e)
        {
            ShowCheckOptionsDialog("MAC", async (clientsToCheck) =>
            {
                foreach (var client in clientsToCheck)
                {
                    try
                    {
                        var ping = new Ping();
                        var reply = await ping.SendPingAsync(client.IP, 1000);

                        if (reply.Status != IPStatus.Success)
                        {
                            UpdateGridRow(client.Name, client.IP, "Unreachable");
                            continue;
                        }

                        using (var sshClient = new SshClient(client.IP, "root", "123456"))
                        {
                            await Task.Run(() => sshClient.Connect());

                            var macCommand = "ip link show | grep -i 'link/ether' | awk '{print $2}' | head -n 1";
                            var actualMac = await Task.Run(() =>
                                sshClient.RunCommand(macCommand).Result.Trim().ToUpper());

                            client.ActualMacAddress = actualMac;
                            UpdateGridRowWithMac(client.Name, client.IP, client.MacAddress, actualMac, "MAC Checked");

                            sshClient.Disconnect();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogConnection(client.Name, client.IP, "Check MAC", "Failed", ex.Message);
                        UpdateGridRow(client.Name, client.IP, "MAC Check Failed");
                    }
                }
            });
        }

        private void ShowCheckOptionsDialog(string operation, Func<List<ClientInfo>, Task> checkAction)
        {
            var optionsForm = new Form
            {
                Text = $"Select {operation} Check Options",
                Size = new Size(400, 500),
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                FormBorderStyle = FormBorderStyle.FixedDialog
            };

            var mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(10)
            };

            // Radio buttons for selection mode
            var radioPanel = new Panel { Dock = DockStyle.Fill };
            var allClientsRadio = new RadioButton
            {
                Text = "All Clients",
                Checked = true,
                Location = new Point(10, 10),
                AutoSize = true
            };

            var selectedClientsRadio = new RadioButton
            {
                Text = "Selected Clients",
                Location = new Point(10, 35),
                AutoSize = true
            };

            var categoryClientsRadio = new RadioButton
            {
                Text = "By Category",
                Location = new Point(10, 60),
                AutoSize = true
            };

            radioPanel.Controls.AddRange(new Control[] {
        allClientsRadio, selectedClientsRadio, categoryClientsRadio
    });

            // Category selection panel
            var categoryPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 5,
                Visible = false
            };

            // Create category dropdowns
            var monitorDropDown = CreateCategoryDropDown("Monitor");
            var typeDropDown = CreateCategoryDropDown("Type");
            var groupDropDown = CreateCategoryDropDown("Group");
            var levelDropDown = CreateCategoryDropDown("Level");
            var zoneDropDown = CreateCategoryDropDown("Zone");

            AddLabelAndDropDown(categoryPanel, "Monitor:", monitorDropDown, 0);
            AddLabelAndDropDown(categoryPanel, "Type:", typeDropDown, 1);
            AddLabelAndDropDown(categoryPanel, "Group:", groupDropDown, 2);
            AddLabelAndDropDown(categoryPanel, "Level:", levelDropDown, 3);
            AddLabelAndDropDown(categoryPanel, "Zone:", zoneDropDown, 4);

            // Populate dropdowns with unique values
            PopulateCategoryDropDowns(new[] { monitorDropDown, typeDropDown, groupDropDown, levelDropDown, zoneDropDown });

            // Client list for selected/filtered clients
            var clientsList = new ListBox
            {
                Dock = DockStyle.Fill,
                SelectionMode = SelectionMode.MultiExtended,
                Visible = false
            };

            // Progress panel
            var progressPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50
            };

            var progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Continuous,
                Height = 20,
                Width = 350,
                Location = new Point(10, 5)
            };

            var progressLabel = new Label
            {
                AutoSize = true,
                Location = new Point(10, 30)
            };

            progressPanel.Controls.AddRange(new Control[] { progressBar, progressLabel });

            // Buttons
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 40
            };

            var startButton = new Button
            {
                Text = "Start Check",
                Width = 100,
                Height = 30
            };

            var cancelButton = new Button
            {
                Text = "Cancel",
                Width = 100,
                Height = 30
            };

            buttonPanel.Controls.AddRange(new Control[] { cancelButton, startButton });

            // Event handlers for radio buttons
            categoryClientsRadio.CheckedChanged += (s, e) =>
            {
                categoryPanel.Visible = categoryClientsRadio.Checked;
                clientsList.Visible = false;
                UpdateClientsList();
            };

            selectedClientsRadio.CheckedChanged += (s, e) =>
            {
                clientsList.Visible = selectedClientsRadio.Checked;
                categoryPanel.Visible = false;

                // Populate with selected clients from main grid
                clientsList.Items.Clear();
                foreach (DataGridViewRow row in gridView.SelectedRows)
                {
                    clientsList.Items.Add($"{row.Cells[0].Value} ({row.Cells[1].Value})");
                }
            };

            allClientsRadio.CheckedChanged += (s, e) =>
            {
                categoryPanel.Visible = false;
                clientsList.Visible = false;
            };

            // Event handlers for category dropdowns
            EventHandler categoryChanged = (s, e) => UpdateClientsList();
            monitorDropDown.SelectedIndexChanged += categoryChanged;
            typeDropDown.SelectedIndexChanged += categoryChanged;
            groupDropDown.SelectedIndexChanged += categoryChanged;
            levelDropDown.SelectedIndexChanged += categoryChanged;
            zoneDropDown.SelectedIndexChanged += categoryChanged;

            // Start button click handler
            startButton.Click += async (s, e) =>
            {
                List<ClientInfo> clientsToCheck = new List<ClientInfo>();

                if (allClientsRadio.Checked)
                {
                    clientsToCheck = clients;
                }
                else if (selectedClientsRadio.Checked)
                {
                    var selectedIndices = gridView.SelectedRows.Cast<DataGridViewRow>()
                        .Select(r => gridView.Rows.IndexOf(r));
                    clientsToCheck = selectedIndices.Select(i => clients[i]).ToList();
                }
                else if (categoryClientsRadio.Checked)
                {
                    clientsToCheck = clients.Where(c =>
                        (monitorDropDown.SelectedIndex == 0 || c.Monitor == monitorDropDown.Text) &&
                        (typeDropDown.SelectedIndex == 0 || c.Type == typeDropDown.Text) &&
                        (groupDropDown.SelectedIndex == 0 || c.Group == groupDropDown.Text) &&
                        (levelDropDown.SelectedIndex == 0 || c.Level == levelDropDown.Text) &&
                        (zoneDropDown.SelectedIndex == 0 || c.Zone == zoneDropDown.Text)
                    ).ToList();
                }

                if (!clientsToCheck.Any())
                {
                    MessageBox.Show("No clients selected for checking.", "Warning",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                progressBar.Maximum = clientsToCheck.Count;
                progressBar.Value = 0;
                startButton.Enabled = false;
                progressLabel.Text = $"Processing 0/{clientsToCheck.Count} clients...";

                try
                {
                    await checkAction(clientsToCheck);
                    MessageBox.Show($"{operation} check completed successfully!", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    optionsForm.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error during {operation} check: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    startButton.Enabled = true;
                }
            };

            cancelButton.Click += (s, e) => optionsForm.Close();

            // Add controls to form
            mainLayout.Controls.Add(radioPanel, 0, 0);
            mainLayout.Controls.Add(categoryPanel, 0, 1);
            mainLayout.Controls.Add(clientsList, 0, 2);
            mainLayout.Controls.Add(progressPanel, 0, 3);
            optionsForm.Controls.Add(mainLayout);
            optionsForm.Controls.Add(buttonPanel);

            void UpdateClientsList()
            {
                if (!categoryClientsRadio.Checked) return;

                var filteredClients = clients.Where(c =>
                    (monitorDropDown.SelectedIndex == 0 || c.Monitor == monitorDropDown.Text) &&
                    (typeDropDown.SelectedIndex == 0 || c.Type == typeDropDown.Text) &&
                    (groupDropDown.SelectedIndex == 0 || c.Group == groupDropDown.Text) &&
                    (levelDropDown.SelectedIndex == 0 || c.Level == levelDropDown.Text) &&
                    (zoneDropDown.SelectedIndex == 0 || c.Zone == zoneDropDown.Text)
                ).ToList();

                clientsList.Items.Clear();
                foreach (var client in filteredClients)
                {
                    clientsList.Items.Add($"{client.Name} ({client.IP})");
                }
            }

            optionsForm.ShowDialog();
        }

        private ComboBox CreateCategoryDropDown(string category)
        {
            return new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(3)
            };
        }

        private void AddLabelAndDropDown(TableLayoutPanel panel, string labelText, ComboBox dropDown, int row)
        {
            panel.Controls.Add(new Label { Text = labelText, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleRight }, 0, row);
            panel.Controls.Add(dropDown, 1, row);
        }

        private void PopulateCategoryDropDowns(ComboBox[] dropDowns)
        {
            // Helper function to populate a dropdown with unique values
            void PopulateDropDown(ComboBox dropDown, Func<ClientInfo, string> selector)
            {
                var values = clients.Select(selector)
                    .Where(v => !string.IsNullOrEmpty(v))
                    .Distinct()
                    .OrderBy(v => v)
                    .ToList();

                dropDown.Items.Clear();
                dropDown.Items.Add("All");
                dropDown.Items.AddRange(values.Cast<object>().ToArray());
                dropDown.SelectedIndex = 0;
            }

            PopulateDropDown(dropDowns[0], c => c.Monitor);
            PopulateDropDown(dropDowns[1], c => c.Type);
            PopulateDropDown(dropDowns[2], c => c.Group);
            PopulateDropDown(dropDowns[3], c => c.Level);
            PopulateDropDown(dropDowns[4], c => c.Zone);
        }

        private void AddLoadButton()
        {
            // Create a new panel for the load button
            var loadButtonPanel = new Panel
            {
                Height = 40,
                Dock = DockStyle.Bottom
            };

            loadClientsButton = new Button
            {
                Text = "Load/Refresh Clients from Excel",
                Width = 200,
                Location = new Point(10, 5),
                Height = 30
            };
            loadClientsButton.Click += LoadClientsButton_Click;

            loadButtonPanel.Controls.Add(loadClientsButton);

            // Add the panel to mainLayout at position 0 (top)
            mainLayout.Controls.Add(loadButtonPanel, 0, 4);
        }

        private async void LoadClientsButton_Click(object sender, EventArgs e)
        {
            try
            {
                var result = MessageBox.Show(
                    "Do you want to reload clients from Excel file?\nThis will refresh all client data.",
                    "Reload Clients",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    loadClientsButton.Enabled = false;
                    statusLabel.Text = "Loading clients from Excel...";
                    statusLabel.ForeColor = Color.Blue;

                    await Task.Run(() => LoadIPsFromExcel(EXCEL_FILE_PATH));
                    SaveClientsToCache();

                    statusLabel.Text = $"Loaded {clients.Count} clients";
                    statusLabel.ForeColor = Color.Green;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Error loading clients";
                statusLabel.ForeColor = Color.Red;
            }
            finally
            {
                loadClientsButton.Enabled = true;
            }
        }

        private void LoadCachedClients()
        {
            try
            {
                if (File.Exists(CACHE_FILE_PATH))
                {
                    statusLabel.Text = "Loading cached client data...";
                    statusLabel.ForeColor = Color.Blue;

                    var serializer = new XmlSerializer(typeof(List<ClientInfo>));
                    using (var stream = new FileStream(CACHE_FILE_PATH, FileMode.Open))
                    {
                        clients = (List<ClientInfo>)serializer.Deserialize(stream);
                    }

                    // Update the grid with cached data
                    UpdateGridWithClients();
                    UpdateFilterOptions();

                    statusLabel.Text = $"Loaded {clients.Count} clients from cache";
                    statusLabel.ForeColor = Color.Green;
                    LogLoadingSummary();
                }
                else
                {
                    statusLabel.Text = "No cached data found. Please load clients from Excel.";
                    statusLabel.ForeColor = Color.Blue;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading cache: {ex}");
                statusLabel.Text = "Error loading cached data. Please load clients from Excel.";
                statusLabel.ForeColor = Color.Red;
            }
        }

        private void SaveClientsToCache()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(List<ClientInfo>));
                using (var stream = new FileStream(CACHE_FILE_PATH, FileMode.Create))
                {
                    serializer.Serialize(stream, clients);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving cache: {ex}");
            }
        }

        private void UpdateGridWithClients()
        {
            if (gridView.InvokeRequired)
            {
                gridView.Invoke(new Action(UpdateGridWithClients));
                return;
            }

            gridView.Rows.Clear();
            foreach (var client in clients)
            {
                AddRowToGrid(client);
            }
        }

        // Add menu items for cache management
        private void AddCacheManagementMenu()
        {
            var menuStrip = new MenuStrip();
            var fileMenu = new ToolStripMenuItem("File");
            var clearCacheItem = new ToolStripMenuItem("Clear Cache", null, (s, e) =>
            {
                var result = MessageBox.Show(
                    "Are you sure you want to clear the cached client data?",
                    "Clear Cache",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        if (File.Exists(CACHE_FILE_PATH))
                        {
                            File.Delete(CACHE_FILE_PATH);
                        }
                        clients.Clear();
                        gridView.Rows.Clear();
                        UpdateFilterOptions();
                        statusLabel.Text = "Cache cleared. Please load clients from Excel.";
                        statusLabel.ForeColor = Color.Blue;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error clearing cache: {ex.Message}", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            });

            fileMenu.DropDownItems.Add(clearCacheItem);
            menuStrip.Items.Add(fileMenu);
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);
        }

        private void ModifyButton_Click(object sender, EventArgs e)
        {
            if (gridView.SelectedRows.Count != 1) return;

            var selectedRow = gridView.SelectedRows[0];
            var modifyForm = new Form
            {
                Text = "Modify Client Information",
                Size = new Size(500, 500),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                ColumnCount = 2,
                RowCount = 9
            };

            // Add fields for modification
            var fields = new Dictionary<string, (TextBox TextBox, string OriginalValue)>
    {
        { "Name", (new TextBox(), selectedRow.Cells["Name"].Value.ToString()) },
        { "IP", (new TextBox(), selectedRow.Cells["IP"].Value.ToString()) },
        { "MAC", (new TextBox(), selectedRow.Cells["MAC"].Value.ToString()) },
        { "Monitor", (new TextBox(), selectedRow.Cells["Monitor"].Value.ToString()) },
        { "Type", (new TextBox(), selectedRow.Cells["Type"].Value.ToString()) },
        { "Group", (new TextBox(), selectedRow.Cells["Group"].Value.ToString()) },
        { "Level", (new TextBox(), selectedRow.Cells["Level"].Value.ToString()) },
        { "Zone", (new TextBox(), selectedRow.Cells["Zone"].Value.ToString()) }
    };

            int row = 0;
            foreach (var field in fields)
            {
                layout.Controls.Add(new Label { Text = field.Key, Dock = DockStyle.Fill }, 0, row);
                field.Value.TextBox.Text = field.Value.OriginalValue;
                field.Value.TextBox.Dock = DockStyle.Fill;
                layout.Controls.Add(field.Value.TextBox, 1, row);
                row++;
            }

            // Add buttons
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft
            };

            var saveButton = new Button { Text = "Save", Width = 100 };
            var cancelButton = new Button { Text = "Cancel", Width = 100 };

            saveButton.Click += async (s, ev) =>
            {
                try
                {
                    var result = MessageBox.Show(
                        "Save changes to Excel file?",
                        "Confirm Save",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                    {
                        modifyForm.Close();
                        return;
                    }

                    Cursor.Current = Cursors.WaitCursor;
                    loadClientsButton.Enabled = false;
                    modifyForm.Enabled = false;

                    // Create backup first
                    string backupFile = Path.Combine(
                        Path.GetDirectoryName(EXCEL_FILE_PATH),
                        Path.GetFileNameWithoutExtension(EXCEL_FILE_PATH) + "_backup.xlsx"
                    );

                    if (File.Exists(backupFile))
                        File.Delete(backupFile);

                    File.Copy(EXCEL_FILE_PATH, backupFile);

                    await Task.Run(() =>
                    {
                        try
                        {
                            // Create a new workbook to copy data to
                            using (var sourceWorkbook = new XLWorkbook(EXCEL_FILE_PATH))
                            using (var newWorkbook = new XLWorkbook())
                            {
                                var sourceWorksheet = sourceWorkbook.Worksheet(EXCEL_SHEET_NAME);
                                var newWorksheet = newWorkbook.Worksheets.Add(EXCEL_SHEET_NAME);

                                // Copy the used range from source to new worksheet
                                var usedRange = sourceWorksheet.RangeUsed();
                                if (usedRange != null)
                                {
                                    // Copy content and formatting
                                    var firstCell = newWorksheet.Cell(1, 1);
                                    usedRange.CopyTo(firstCell);
                                }

                                // Find and update our specific row
                                var rows = newWorksheet.RowsUsed();
                                bool foundRow = false;

                                foreach (var xlRow in rows)
                                {
                                    string nameCell = xlRow.Cell("A").GetString().Trim();
                                    string ipCell = xlRow.Cell("B").GetString().Trim();

                                    if (nameCell == fields["Name"].OriginalValue &&
                                        ipCell == fields["IP"].OriginalValue)
                                    {
                                        foundRow = true;
                                        xlRow.Cell("A").Value = fields["Name"].TextBox.Text;
                                        xlRow.Cell("B").Value = fields["IP"].TextBox.Text;
                                        xlRow.Cell("C").Value = fields["MAC"].TextBox.Text;
                                        xlRow.Cell("D").Value = fields["Monitor"].TextBox.Text;
                                        xlRow.Cell("E").Value = fields["Type"].TextBox.Text;
                                        xlRow.Cell("F").Value = fields["Group"].TextBox.Text;
                                        xlRow.Cell("G").Value = fields["Level"].TextBox.Text;
                                        xlRow.Cell("H").Value = fields["Zone"].TextBox.Text;
                                        break;
                                    }
                                }

                                if (!foundRow)
                                    throw new Exception("Could not find the row to update.");

                                // Save to a new temporary file
                                string tempFile = Path.Combine(
                                    Path.GetDirectoryName(EXCEL_FILE_PATH),
                                    $"temp_{DateTime.Now:yyyyMMddHHmmss}.xlsx"
                                );

                                newWorkbook.SaveAs(tempFile);

                                // Close all workbooks
                                newWorkbook.Dispose();
                                sourceWorkbook.Dispose();

                                // Small delay to ensure files are released
                                Thread.Sleep(500);

                                // Replace the original file
                                if (File.Exists(EXCEL_FILE_PATH))
                                    File.Delete(EXCEL_FILE_PATH);

                                File.Move(tempFile, EXCEL_FILE_PATH);

                                // Update client object in the list
                                var client = clients.FirstOrDefault(c =>
                                    c.Name == fields["Name"].OriginalValue &&
                                    c.IP == fields["IP"].OriginalValue);

                                if (client != null)
                                {
                                    client.Name = fields["Name"].TextBox.Text;
                                    client.IP = fields["IP"].TextBox.Text;
                                    client.MacAddress = fields["MAC"].TextBox.Text;
                                    client.Monitor = fields["Monitor"].TextBox.Text;
                                    client.Type = fields["Type"].TextBox.Text;
                                    client.Group = fields["Group"].TextBox.Text;
                                    client.Level = fields["Level"].TextBox.Text;
                                    client.Zone = fields["Zone"].TextBox.Text;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // If anything fails, restore from backup
                            if (File.Exists(backupFile))
                            {
                                if (File.Exists(EXCEL_FILE_PATH))
                                    File.Delete(EXCEL_FILE_PATH);
                                File.Copy(backupFile, EXCEL_FILE_PATH);
                            }
                            throw;
                        }
                        finally
                        {
                            // Clean up backup file
                            if (File.Exists(backupFile))
                                File.Delete(backupFile);
                        }
                    });

                    // Update grid view
                    selectedRow.Cells["Name"].Value = fields["Name"].TextBox.Text;
                    selectedRow.Cells["IP"].Value = fields["IP"].TextBox.Text;
                    selectedRow.Cells["MAC"].Value = fields["MAC"].TextBox.Text;
                    selectedRow.Cells["Monitor"].Value = fields["Monitor"].TextBox.Text;
                    selectedRow.Cells["Type"].Value = fields["Type"].TextBox.Text;
                    selectedRow.Cells["Group"].Value = fields["Group"].TextBox.Text;
                    selectedRow.Cells["Level"].Value = fields["Level"].TextBox.Text;
                    selectedRow.Cells["Zone"].Value = fields["Zone"].TextBox.Text;

                    // Update the cache
                    SaveClientsToCache();

                    MessageBox.Show("Changes saved successfully!", "Success",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    modifyForm.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving changes: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    loadClientsButton.Enabled = true;
                    modifyForm.Enabled = true;
                    Cursor.Current = Cursors.Default;
                }
            };

            cancelButton.Click += (s, ev) => modifyForm.Close();

            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(saveButton);
            layout.Controls.Add(buttonPanel, 0, 8);
            layout.SetColumnSpan(buttonPanel, 2);

            modifyForm.Controls.Add(layout);
            modifyForm.ShowDialog();
        }
    }

}