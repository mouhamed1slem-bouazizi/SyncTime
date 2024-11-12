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

namespace SyncTime
{
    public class ClientInfo
    {
        public string Name { get; set; }
        public string IP { get; set; }
    }

    public partial class Form1 : Form
    {
        private const string EXCEL_FILES_DIRECTORY = @"C:\IP";
        private readonly Dictionary<string, string> excelFiles = new Dictionary<string, string>
        {
            { "IntGates", "intGates.xlsx" },
            { "DomGates", "domGates.xlsx" },
            { "Carousels", "carousel.xlsx" },
            { "PUB_Arr", "PUB_arr.xlsx" },
            { "PUB_Boarding", "PUB_boa.xlsx" },
            { "PUB_Dep", "PUB_dep.xlsx" },
            { "Line_A", "line_a.xlsx" },
            { "Line_B", "line_b.xlsx" },
            { "Line_C", "line_c.xlsx" },
            { "Line_D", "line_d.xlsx" },
            { "Line_E", "line_e.xlsx" },
            { "Line_F", "line_f.xlsx" },
            { "Line_G", "line_g.xlsx" },
            { "VWall", "vw.xlsx" },
            { "GVIP", "gvip.xlsx" }
        };

        private List<ClientInfo> clients = new List<ClientInfo>();
        private DataGridView gridView;
        private Button syncButton;
        private Button executeCodeButton;
        private Label statusLabel;
        private TableLayoutPanel mainLayout;
        private ComboBox fileSelector;
        private DataGridViewRow selectedRow;

        public Form1()
        {
            InitializeComponent();
            InitializeUI();
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
                AutoSize = true
            };

            // Configure row styles
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // File selector
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));  // Status label
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Grid
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));  // Button

            this.Controls.Add(mainLayout);

            // File Selector Panel
            Panel fileSelectorPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 5, 0, 5)
            };
            mainLayout.Controls.Add(fileSelectorPanel, 0, 0);

            // File Selector ComboBox
            fileSelector = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 200,
                Anchor = AnchorStyles.None
            };

            // Add items to ComboBox
            foreach (var file in excelFiles.Keys)
            {
                fileSelector.Items.Add(file);
            }

            fileSelector.SelectedIndexChanged += FileSelector_SelectedIndexChanged;

            // Center the ComboBox in its panel
            fileSelector.Location = new Point(
                (fileSelectorPanel.Width - fileSelector.Width) / 2,
                (fileSelectorPanel.Height - fileSelector.Height) / 2
            );
            fileSelectorPanel.Controls.Add(fileSelector);

            // Handle panel resize to keep ComboBox centered
            fileSelectorPanel.Resize += (s, e) =>
            {
                fileSelector.Location = new Point(
                    (fileSelectorPanel.Width - fileSelector.Width) / 2,
                    (fileSelectorPanel.Height - fileSelector.Height) / 2
                );
            };

            // Status Label
            statusLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.Blue,
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 5, 0, 5)
            };
            mainLayout.Controls.Add(statusLabel, 0, 1);

            // Grid View Panel
            Panel gridPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 5, 0, 5)
            };
            mainLayout.Controls.Add(gridPanel, 0, 2);

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
            gridView.Columns.Add("Status", "Status");
            gridView.Columns.Add("CurrentTime", "Current Time");
            gridView.Columns.Add("LastSync", "Last Sync Time");

            // Set column weights
            gridView.Columns[0].FillWeight = 20;
            gridView.Columns[1].FillWeight = 20;
            gridView.Columns[2].FillWeight = 15;
            gridView.Columns[3].FillWeight = 22.5F;
            gridView.Columns[4].FillWeight = 22.5F;

            gridPanel.Controls.Add(gridView);

            // Add selection change handler for the grid
            gridView.CellClick += GridView_CellClick;

            // Button Panel
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill
            };
            mainLayout.Controls.Add(buttonPanel, 0, 3);

            // Create a TableLayoutPanel for the buttons
            TableLayoutPanel buttonsLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
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

            // Add buttons to the layout
            buttonsLayout.Controls.Add(syncButton, 0, 0);
            buttonsLayout.Controls.Add(executeCodeButton, 1, 0);

            // Center the buttons in their cells
            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            // Style the grid
            gridView.EnableHeadersVisualStyles = false;
            gridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 66, 91);
            gridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            gridView.ColumnHeadersDefaultCellStyle.Font = new Font(gridView.Font, FontStyle.Bold);
            gridView.ColumnHeadersHeight = 35;
            gridView.RowTemplate.Height = 30;
            gridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250);

            // Set initial status
            statusLabel.Text = "Please select a FIDS system";
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
                UpdateGridRow(clientName, clientIP, "Connecting...");
                var ping = new Ping();
                var reply = await ping.SendPingAsync(clientIP, 1000);

                if (reply.Status != IPStatus.Success)
                {
                    UpdateGridRow(clientName, clientIP, "Unreachable");
                    MessageBox.Show("Client is unreachable.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                sshClient = new SshClient(clientIP, "root", "123456");
                await Task.Run(() => sshClient.Connect());
                isConnected = true;
                UpdateGridRow(clientName, clientIP, "Connected");
                resultsBox.AppendText("Connected to " + clientIP + "\r\n");
            }
            catch (Exception ex)
            {
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
        private async void FileSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (fileSelector.SelectedItem == null) return;

            string selectedSystem = fileSelector.SelectedItem.ToString();
            string fileName = excelFiles[selectedSystem];
            string filePath = Path.Combine(EXCEL_FILES_DIRECTORY, fileName);

            try
            {
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"Excel file not found: {fileName}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    statusLabel.Text = "Excel file not found";
                    statusLabel.ForeColor = Color.Red;
                    return;
                }

                statusLabel.Text = $"Loading {selectedSystem} clients...";
                statusLabel.ForeColor = Color.Blue;

                await Task.Run(() => LoadIPsFromExcel(filePath));

                if (clients.Any())
                {
                    syncButton.Enabled = true;
                    await RefreshClientTimes();
                    statusLabel.Text = $"Loaded {clients.Count} clients from {selectedSystem}";
                    statusLabel.ForeColor = Color.Green;
                }
                else
                {
                    statusLabel.Text = $"No clients found in {selectedSystem} file";
                    statusLabel.ForeColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading clients: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Error loading clients";
                statusLabel.ForeColor = Color.Red;
            }
        }

        private void LoadIPsFromExcel(string filePath)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    clients.Clear();

                    if (gridView.InvokeRequired)
                    {
                        gridView.Invoke(new Action(() => gridView.Rows.Clear()));
                    }
                    else
                    {
                        gridView.Rows.Clear();
                    }

                    // Skip header row if exists
                    reader.Read();

                    while (reader.Read())
                    {
                        string name = reader.GetValue(0)?.ToString();
                        string ip = reader.GetValue(1)?.ToString();

                        if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(ip))
                        {
                            clients.Add(new ClientInfo { Name = name, IP = ip });

                            if (gridView.InvokeRequired)
                            {
                                gridView.Invoke(new Action(() =>
                                    gridView.Rows.Add(name, ip, "Not checked", "-", "-")));
                            }
                            else
                            {
                                gridView.Rows.Add(name, ip, "Not checked", "-", "-");
                            }
                        }
                    }
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

        private async Task SynchronizeTime()
        {
            string currentTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            List<string> unreachableClients = new List<string>();

            await Parallel.ForEachAsync(clients, async (client, token) =>
            {
                try
                {
                    // Check if client is reachable first
                    var ping = new Ping();
                    var reply = await ping.SendPingAsync(client.IP, 1000); // 1 second timeout

                    if (reply.Status != IPStatus.Success)
                    {
                        UpdateGridRow(client.Name, client.IP, "Unreachable");
                        unreachableClients.Add($"{client.Name} ({client.IP})");
                        return; // Skip this client
                    }

                    using (var sshClient = new SshClient(client.IP, "root", "123456"))
                    {
                        UpdateGridRow(client.Name, client.IP, "Connecting...");
                        await Task.Run(() => sshClient.Connect());

                        UpdateGridRow(client.Name, client.IP, "Setting time...");
                        string dateCommand = $"date -s \"{currentTime}\"";
                        await Task.Run(() => sshClient.RunCommand(dateCommand));
                        await Task.Run(() => sshClient.RunCommand("hwclock --systohc"));

                        // Verify time was set correctly
                        var timeOutput = await Task.Run(() =>
                            sshClient.RunCommand("date '+%Y-%m-%d %H:%M:%S'").Result.Trim());

                        sshClient.Disconnect();
                        UpdateGridRow(client.Name, client.IP, "Synchronized", timeOutput);
                    }
                }
                catch (Exception)
                {
                    UpdateGridRow(client.Name, client.IP, "Connection Failed");
                    unreachableClients.Add($"{client.Name} ({client.IP})");
                }
            });

            // Show summary of unreachable clients
            if (unreachableClients.Any())
            {
                string message = "The following clients were unreachable or failed:\n\n" +
                               string.Join("\n", unreachableClients);
                MessageBox.Show(message, "Unreachable Clients", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void UpdateGridRow(string name, string ip, string status, string time = "-")
        {
            if (gridView.InvokeRequired)
            {
                gridView.Invoke(new Action(() => UpdateGridRow(name, ip, status, time)));
                return;
            }

            foreach (DataGridViewRow row in gridView.Rows)
            {
                if (row.Cells[0].Value.ToString() == name && row.Cells[1].Value.ToString() == ip)
                {
                    row.Cells[2].Value = status;
                    if (time != "-")
                    {
                        row.Cells[3].Value = time;
                        if (status == "Synchronized")
                        {
                            row.Cells[4].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }

                    // Color code the status
                    switch (status)
                    {
                        case "Synchronized":
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                            break;
                        case "Unreachable":
                        case "Connection Failed":
                            row.DefaultCellStyle.BackColor = Color.LightPink;
                            break;
                        case "Connecting...":
                        case "Setting time...":
                            row.DefaultCellStyle.BackColor = Color.LightYellow;
                            break;
                        default:
                            row.DefaultCellStyle.BackColor = Color.White;
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
                        var timeOutput = await Task.Run(() =>
                            sshClient.RunCommand("date '+%Y-%m-%d %H:%M:%S'").Result.Trim());
                        sshClient.Disconnect();

                        UpdateGridRow(client.Name, client.IP, "Connected", timeOutput);
                    }
                }
                catch (Exception)
                {
                    UpdateGridRow(client.Name, client.IP, "Connection Failed");
                }
            }
        }
    }
}