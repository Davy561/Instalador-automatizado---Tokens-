using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Security;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Principal;

namespace Script_de_Cache
{
    public class DriverInfo
    {
        public string Name { get; set; }
        public string SilentArgs { get; set; }
        public string Description { get; set; }
        public string FileName { get; set; }
        public bool IsSelected { get; set; } = true;
    }

    public partial class Form1 : Form
    {
        // ===== CREDENCIAIS DO AD =====
        private readonly Dictionary<string, string> AD_ACCOUNTS = new Dictionary<string, string>
        {
            { "dpcorreir", "senha" }
        };

        // ===== CREDENCIAIS LOCAIS =====
        private readonly Dictionary<string, string> LOCAL_ACCOUNTS = new Dictionary<string, string>
        {
            { "adminvix", "senha" },
            { "Administrador", "senha" }
        };

        private const string DOMAIN = "PMV";
        private const string DOMAIN_CONTROLLER = "pmv.local";

        private string chosenUsername = null;
        private SecureString chosenPassword = null;
        private bool isDomainReachable = false;

        private string downloadDir;
        private List<DriverInfo> allDrivers;

        // Componentes da UI de Seleção
        private Panel selectionPanel;
        private Panel buttonsPanel;
        private Button btnInstall;
        private Button btnSelectAll;
        private Button btnDeselectAll;
        private Label lblTitle;
        private List<CheckBox> driverCheckBoxes;

        // Componentes da UI de Progresso
        private Panel progressPanel;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblProgressInfo;

        private bool isInstalling = false;

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            InitializeDirectories();
            CheckDomainAndSelectAccount();
            InitializeDrivers();
            CreateSelectionInterface();
        }

        private void CheckDomainAndSelectAccount()
        {
            isDomainReachable = IsDomainControllerReachable(DOMAIN_CONTROLLER);

            // Escolhe primeira conta AD
            if (AD_ACCOUNTS.Count > 0)
            {
                var firstAccount = AD_ACCOUNTS.First();
                chosenUsername = firstAccount.Key;

                chosenPassword = new SecureString();
                foreach (char c in firstAccount.Value)
                {
                    chosenPassword.AppendChar(c);
                }
                chosenPassword.MakeReadOnly();
            }
        }

        private bool IsDomainControllerReachable(string dcName)
        {
            try
            {
                System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();
                var reply = ping.Send(dcName, 3000);
                return reply.Status == System.Net.NetworkInformation.IPStatus.Success;
            }
            catch
            {
                return false;
            }
        }

        private void InitializeCustomComponents()
        {
            this.Text = "Instalador de Drivers - Selecione os Drivers";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.BackColor = Color.FromArgb(245, 245, 245);

            // Carregar ícone dos recursos embutidos
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string[] allResources = assembly.GetManifestResourceNames();

                // Procurar pelo recurso do ícone
                string iconResource = allResources.FirstOrDefault(r =>
                    r.EndsWith("R.ico", StringComparison.OrdinalIgnoreCase) ||
                    r.EndsWith(".ico", StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrEmpty(iconResource))
                {
                    using (Stream iconStream = assembly.GetManifestResourceStream(iconResource))
                    {
                        if (iconStream != null)
                        {
                            this.Icon = new Icon(iconStream);
                        }
                    }
                }
                else
                {
                    this.Icon = SystemIcons.Application;
                }
            }
            catch
            {
                this.Icon = SystemIcons.Application;
            }
        }

        private void InitializeDirectories()
        {
            downloadDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads", "token-drivers"
            );

            Directory.CreateDirectory(downloadDir);
        }

        private void InitializeDrivers()
        {
            allDrivers = new List<DriverInfo>
            {
                new DriverInfo
                {
                    Name = "Feitian ePass2003 Setup v1.1.16.330",
                    SilentArgs = "/S",
                    Description = "Instalador completo ePass 2003",
                    FileName = "ePass2003-Setup_v1.1.16.330.exe",
                    IsSelected = false
                },
                new DriverInfo
                {
                    Name = "Instalador DXSafe Middleware v1.0.30",
                    SilentArgs = "/S",
                    Description = "Middleware para Token DX Safe (Soluti/Certisign)",
                    FileName = "Instalador DXSafe Middleware - 1.0.30.exe",
                    IsSelected = false
                },
                new DriverInfo
                {
                    Name = "Safenet 10.7",
                    SilentArgs = "/quiet /norestart",
                    Description = "Assinador digital Safenet",
                    FileName = "Safenet_10.7.msi",
                    IsSelected = false
                },
                new DriverInfo
                {
                    Name = "SafeSign",
                    SilentArgs = "/S",
                    Description = "Assinador digital SafeSign",
                    FileName = "SafeSignIC30124-x64-win-tu-admin.exe",
                    IsSelected = false
                },
                new DriverInfo
                {
                    Name = "GD Starsign",
                    Description = "Assinador digital GD",
                    FileName = "GDsetupStarsignCUTx64.exe",
                    IsSelected = false
                },
                 new DriverInfo
                {
                    Name = "ePass Manager Admin 2003",
                    SilentArgs = "/S",
                    Description = "Ferramenta de gerenciamento ePass 2003",
                    FileName = "ePassManagerAdm_2003.exe",
                    IsSelected = false
                },
            };
        }

        private void CreateSelectionInterface()
        {
            // ===== PAINEL PRINCIPAL DE SELEÇÃO =====
            selectionPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(20),
                Visible = true
            };

            // Título
            lblTitle = new Label
            {
                Text = "Selecione o modelo do Token que deseja instalar:",
                Font = new Font("Segoe UI", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(50, 50, 50),
                AutoSize = true,
                Location = new Point(20, 20)
            };
            selectionPanel.Controls.Add(lblTitle);

            // Painel com scroll para os checkboxes
            Panel scrollPanel = new Panel
            {
                Location = new Point(20, 60),
                Size = new Size(740, 380),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                AutoScroll = true
            };

            driverCheckBoxes = new List<CheckBox>();
            int yPos = 10;

            for (int i = 0; i < allDrivers.Count; i++)
            {
                var driver = allDrivers[i];

                Panel driverPanel = new Panel
                {
                    Location = new Point(10, yPos),
                    Size = new Size(700, 70),
                    BackColor = i % 2 == 0 ? Color.FromArgb(250, 250, 250) : Color.White
                };

                CheckBox chk = new CheckBox
                {
                    Location = new Point(10, 10),
                    Size = new Size(680, 25),
                    Font = new Font("Segoe UI", 11, FontStyle.Bold),
                    ForeColor = Color.FromArgb(50, 50, 50),
                    Checked = driver.IsSelected,
                    Text = driver.Name,
                    Tag = driver
                };
                chk.CheckedChanged += (s, e) => { driver.IsSelected = chk.Checked; };
                driverCheckBoxes.Add(chk);
                driverPanel.Controls.Add(chk);

                Label lblDesc = new Label
                {
                    Location = new Point(30, 40),
                    Size = new Size(650, 20),
                    Font = new Font("Segoe UI", 9, FontStyle.Italic),
                    ForeColor = Color.FromArgb(120, 120, 120),
                    Text = driver.Description
                };
                driverPanel.Controls.Add(lblDesc);

                scrollPanel.Controls.Add(driverPanel);
                yPos += 75;
            }

            selectionPanel.Controls.Add(scrollPanel);

            // ===== PAINEL DE BOTÕES =====
            buttonsPanel = new Panel
            {
                Location = new Point(20, 450),
                Size = new Size(740, 80),
                BackColor = Color.FromArgb(245, 245, 245)
            };

            btnSelectAll = new Button
            {
                Location = new Point(0, 10),
                Size = new Size(150, 40),
                Text = "✓ Selecionar Todos",
                Font = new Font("Segoe UI", 10),
                BackColor = Color.FromArgb(100, 150, 200),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnSelectAll.FlatAppearance.BorderSize = 0;
            btnSelectAll.Click += BtnSelectAll_Click;
            buttonsPanel.Controls.Add(btnSelectAll);

            btnDeselectAll = new Button
            {
                Location = new Point(160, 10),
                Size = new Size(150, 40),
                Text = "✗ Desmarcar Todos",
                Font = new Font("Segoe UI", 10),
                BackColor = Color.FromArgb(150, 150, 150),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnDeselectAll.FlatAppearance.BorderSize = 0;
            btnDeselectAll.Click += BtnDeselectAll_Click;
            buttonsPanel.Controls.Add(btnDeselectAll);

            btnInstall = new Button
            {
                Location = new Point(540, 10),
                Size = new Size(200, 50),
                Text = "🚀 INSTALAR DRIVER(S)",
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                BackColor = Color.FromArgb(76, 175, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            buttonsPanel.Controls.Add(btnInstall);

            selectionPanel.Controls.Add(buttonsPanel);

            // ===== PAINEL DE PROGRESSO =====
            progressPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(245, 245, 245),
                Padding = new Padding(40),
                Visible = false
            };

            lblStatus = new Label
            {
                Text = "Preparando instalação...",
                Font = new Font("Segoe UI", 12),
                ForeColor = Color.FromArgb(100, 100, 100),
                AutoSize = true,
                Location = new Point(40, 200)
            };
            progressPanel.Controls.Add(lblStatus);

            progressBar = new ProgressBar
            {
                Location = new Point(40, 240),
                Size = new Size(700, 40),
                Style = ProgressBarStyle.Continuous,
                BackColor = Color.FromArgb(220, 220, 220),
                ForeColor = Color.FromArgb(76, 175, 80)
            };
            progressPanel.Controls.Add(progressBar);

            lblProgressInfo = new Label
            {
                Text = "Aguarde...",
                Font = new Font("Segoe UI", 10, FontStyle.Italic),
                ForeColor = Color.FromArgb(150, 150, 150),
                AutoSize = true,
                Location = new Point(40, 290)
            };
            progressPanel.Controls.Add(lblProgressInfo);

            this.Controls.Add(selectionPanel);
            this.Controls.Add(progressPanel);
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            foreach (var chk in driverCheckBoxes)
            {
                chk.Checked = true;
            }
        }

        private void BtnDeselectAll_Click(object sender, EventArgs e)
        {
            foreach (var chk in driverCheckBoxes)
            {
                chk.Checked = false;
            }
        }

        private async void BtnInstall_Click(object sender, EventArgs e)
        {
            if (isInstalling) return;

            var selectedDrivers = allDrivers.Where(d => d.IsSelected).ToList();

            if (selectedDrivers.Count == 0)
            {
                MessageBox.Show(
                    "Por favor, selecione pelo menos um driver para instalar.",
                    "Nenhum driver selecionado",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            var result = MessageBox.Show(
                $"Você selecionou {selectedDrivers.Count} driver(s) para instalação.\n\n" +
                $"Deseja continuar?",
                "Confirmar Instalação",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result != DialogResult.Yes)
                return;

            selectionPanel.Visible = false;
            progressPanel.Visible = true;
            isInstalling = true;

            await StartInstallation(selectedDrivers);
        }

        private async Task StartInstallation(List<DriverInfo> driversToInstall)
        {
            await Task.Delay(500);

            int successCount = 0;
            int failCount = 0;
            List<string> failedDrivers = new List<string>();

            for (int i = 0; i < driversToInstall.Count; i++)
            {
                var driverInfo = driversToInstall[i];
                UpdateProgress(i + 1, driversToInstall.Count);

                bool success = await Task.Run(() => InstallDriver(driverInfo));

                if (success)
                {
                    successCount++;
                }
                else
                {
                    failCount++;
                    failedDrivers.Add(driverInfo.Name);
                }

                await Task.Delay(1000);
            }

            progressBar.Value = 100;
            lblStatus.Text = "Concluído!";

            await Task.Delay(1000);

            string message = $"✓ {successCount} instalado(s)\n";
            if (failCount > 0)
            {
                message += $"✗ {failCount} falharam\n\n";
                message += "Drivers que falharam:\n";
                foreach (var driver in failedDrivers)
                {
                    message += $"  • {driver}\n";
                }
            }
            message += "\n⚠️ Reinicie o computador!";

            MessageBox.Show(message, "Concluído", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Application.Exit();
        }

        private bool InstallDriver(DriverInfo driverInfo)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string[] allResources = assembly.GetManifestResourceNames();

                string resourceName = FindResource(allResources, driverInfo.FileName);
                if (string.IsNullOrEmpty(resourceName))
                {
                    return false;
                }

                Stream resourceStream = assembly.GetManifestResourceStream(resourceName);
                if (resourceStream == null)
                {
                    return false;
                }

                string extension = Path.GetExtension(driverInfo.FileName);
                string tempDir = @"C:\Temp";

                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                string tempPath = Path.Combine(tempDir, $"drv_{Guid.NewGuid().ToString("N").Substring(0, 8)}{extension}");

                using (FileStream fileStream = File.Create(tempPath))
                {
                    resourceStream.CopyTo(fileStream);
                }
                resourceStream.Close();

                System.Threading.Thread.Sleep(300);

                bool installed = InstallWithCredentials(tempPath, driverInfo.SilentArgs);

                try
                {
                    if (File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch { }

                return installed;
            }
            catch
            {
                return false;
            }
        }

        private void KillErrorWindows()
        {
            try
            {
                var processes = Process.GetProcesses();
                foreach (var process in processes)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(process.MainWindowTitle))
                        {
                            string title = process.MainWindowTitle;
                            string titleLower = title.ToLower();

                            bool isErrorWindow =
                                titleLower.Contains("1628") ||
                                (titleLower.Contains("failed") && titleLower.Contains("complete")) ||
                                (titleLower.Contains("falhou") && !titleLower.Contains("wizard")) ||
                                (titleLower.Contains("error") && !titleLower.Contains("wizard") && !titleLower.Contains("installshield"));

                            bool isWizard = titleLower.Contains("installshield wizard") && !titleLower.Contains("1628");

                            if (isErrorWindow && !isWizard)
                            {
                                process.CloseMainWindow();
                                System.Threading.Thread.Sleep(100);
                                if (!process.HasExited)
                                {
                                    process.Kill();
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private bool InstallWithCredentials(string filePath, string silentArgs)
        {
            // ===== TENTATIVA 1: CONTAS DO AD (se domínio acessível) =====
            if (chosenUsername != null && chosenPassword != null && isDomainReachable)
            {
                string[] domainsToTry = { DOMAIN, DOMAIN_CONTROLLER };

                foreach (string domain in domainsToTry)
                {
                    bool success = TryInstallWithDomain(filePath, silentArgs, domain, chosenUsername, chosenPassword);
                    if (success) return true;
                }
            }

            // ===== TENTATIVA 2: CONTAS LOCAIS =====
            if (LOCAL_ACCOUNTS.Count > 0)
            {
                foreach (var localAccount in LOCAL_ACCOUNTS)
                {
                    bool success = TryInstallWithLocalAccount(
                        filePath,
                        silentArgs,
                        localAccount.Key,
                        localAccount.Value
                    );

                    if (success) return true;
                }
            }

            return false;
        }

        private bool TryInstallWithDomain(string filePath, string silentArgs, string domain, string username, SecureString password)
        {
            try
            {
                string extension = Path.GetExtension(filePath).ToLower();

                string fileName = extension == ".msi" ? "msiexec.exe" : filePath;
                string arguments = extension == ".msi"
                    ? $"/i \"{filePath}\" {silentArgs} /qn"
                    : silentArgs;

                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    Domain = domain,
                    UserName = username,
                    Password = password,
                    UseShellExecute = false,
                    LoadUserProfile = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                var proc = Process.Start(psi);

                if (proc == null)
                {
                    return false;
                }

                Task.Run(() =>
                {
                    while (!proc.HasExited)
                    {
                        KillErrorWindows();
                        System.Threading.Thread.Sleep(1000);
                    }
                });

                bool finished = proc.WaitForExit(300000);

                if (!finished)
                {
                    try { proc.Kill(); } catch { }
                    return false;
                }

                int exitCode = proc.ExitCode;

                switch (exitCode)
                {
                    case 0:
                    case 3010:
                    case 1638:
                    case 1641:
                    case -1:
                    case -3:
                    case 1628:
                        return true;
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private bool TryInstallWithLocalAccount(string filePath, string silentArgs, string username, string password)
        {
            try
            {
                string extension = Path.GetExtension(filePath).ToLower();

                string fileName = extension == ".msi" ? "msiexec.exe" : filePath;
                string arguments = extension == ".msi"
                    ? $"/i \"{filePath}\" {silentArgs} /qn"
                    : silentArgs;

                var securePassword = new SecureString();
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }
                securePassword.MakeReadOnly();

                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    Domain = Environment.MachineName,
                    UserName = username,
                    Password = securePassword,
                    UseShellExecute = false,
                    LoadUserProfile = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                var proc = Process.Start(psi);

                if (proc == null)
                {
                    return false;
                }

                Task.Run(() =>
                {
                    while (!proc.HasExited)
                    {
                        KillErrorWindows();
                        System.Threading.Thread.Sleep(1000);
                    }
                });

                bool finished = proc.WaitForExit(300000);

                if (!finished)
                {
                    try { proc.Kill(); } catch { }
                    return false;
                }

                int exitCode = proc.ExitCode;

                switch (exitCode)
                {
                    case 0:
                    case 3010:
                    case 1638:
                    case 1641:
                    case -1:
                    case -3:
                    case 1628:
                        return true;
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private string FindResource(string[] allResources, string fileName)
        {
            var match = allResources.FirstOrDefault(r =>
                r.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));
            if (match != null) return match;

            string normalized = fileName.Replace(" ", "_").Replace("-", "_").Replace(".", "_").ToLower();
            match = allResources.FirstOrDefault(r =>
                r.Replace(" ", "_").Replace("-", "_").Replace(".", "_").ToLower().EndsWith(normalized));
            if (match != null) return match;

            string nameOnly = Path.GetFileNameWithoutExtension(fileName);
            match = allResources.FirstOrDefault(r =>
                r.IndexOf(nameOnly, StringComparison.OrdinalIgnoreCase) >= 0);

            return match;
        }

        private void UpdateProgress(int current, int total)
        {
            int percentage = (current * 100) / total;

            if (progressBar.InvokeRequired)
            {
                progressBar.Invoke(new Action(() =>
                {
                    progressBar.Value = Math.Min(percentage, 100);
                    lblStatus.Text = $"Instalando {current} de {total}...";
                    lblProgressInfo.Text = $"{percentage}%";
                }));
            }
            else
            {
                progressBar.Value = Math.Min(percentage, 100);
                lblStatus.Text = $"Instalando {current} de {total}...";
                lblProgressInfo.Text = $"{percentage}%";
            }
        }
    }
}