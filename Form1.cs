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
        private Label statusLabel;
        private TableLayoutPanel mainLayout;
        private ComboBox fileSelector;

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
                RowCount = 4, // Added one more row for the ComboBox
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
                BorderStyle = BorderStyle.Fixed3D
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

            // Button Panel
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill
            };
            mainLayout.Controls.Add(buttonPanel, 0, 3);

            // Sync Button
            syncButton = new Button
            {
                Text = "Synchronize Time",
                Size = new Size(150, 30),
                Enabled = false,
                Anchor = AnchorStyles.None
            };
            syncButton.Click += SyncButton_Click;

            // Center the button
            syncButton.Location = new Point(
                (buttonPanel.Width - syncButton.Width) / 2,
                (buttonPanel.Height - syncButton.Height) / 2
            );
            buttonPanel.Controls.Add(syncButton);

            // Handle button panel resize
            buttonPanel.Resize += (s, e) =>
            {
                syncButton.Location = new Point(
                    (buttonPanel.Width - syncButton.Width) / 2,
                    (buttonPanel.Height - syncButton.Height) / 2
                );
            };

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