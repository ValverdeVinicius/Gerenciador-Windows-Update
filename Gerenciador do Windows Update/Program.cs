using System;
using System.Drawing;
using System.Windows.Forms;
using System.ServiceProcess;
using System.Diagnostics;
using Microsoft.Win32;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Text;
using System.IO;
using System.Linq;

public static class ControlExtensions
{
    public static void InvokeIfRequired(this Control control, Action action)
    {
        if (control.InvokeRequired)
            control.Invoke(action);
        else
            action();
    }
}

public class WindowsUpdateManager
{
    // Flag para determinar se está no modo gráfico ou modo silencioso
    private bool silentMode;
    private MainForm mainForm;

    public WindowsUpdateManager(MainForm form = null)
    {
        mainForm = form;
        silentMode = (form == null);
    }

    public void LogMessage(string message, bool isError = false, bool forceSilent = false)
    {
        string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        string logMessage = $"[{timestamp}] {message}\n";

        if (forceSilent || silentMode)
        {
            // No modo silencioso, escreve no arquivo de log
            File.AppendAllText("UpdateManager.log", logMessage);
        }
        else if (mainForm != null)
        {
            // No modo gráfico, usa o método de log do formulário
            mainForm.LogToRichTextBox(message, isError);
        }
    }

    public void DisableWindowsUpdate(bool forceSilent = false)
    {
        LogMessage("Iniciando o processo para desabilitar o Windows Update ...", false, forceSilent);

        try
        {
            // Para o serviço do Windows Update primeiro
            using (ServiceController service = new ServiceController("wuauserv"))
            {
                if (service.Status == ServiceControllerStatus.Running)
                {
                    LogMessage("Parando o serviço Windows Update...", false, forceSilent);
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
                    LogMessage("Serviço do Windows Update foi finalizado com sucesso!", false, forceSilent);
                }
            }

            // Desabilita o serviço do Windows Update
            LogMessage("Configurando o tipo de inicialização do serviço para desabilitado...", false, forceSilent);
            RunProcessAsAdmin("cmd.exe", "/c sc config wuauserv start= disabled", forceSilent);

            // Modificações do registro adicionais
            LogMessage("Atualizando as configurações do registro...", false, forceSilent);
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", true))
                {
                    if (key == null)
                    {
                        Registry.LocalMachine.CreateSubKey(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", true);
                        using (RegistryKey newKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", true))
                        {
                            newKey.SetValue("NoAutoUpdate", 1, RegistryValueKind.DWord);
                            newKey.SetValue("AUOptions", 1, RegistryValueKind.DWord);
                        }
                    }
                    else
                    {
                        key.SetValue("NoAutoUpdate", 1, RegistryValueKind.DWord);
                        key.SetValue("AUOptions", 2, RegistryValueKind.DWord); // Notifica antes de fazer o Download
                    }
                }
                LogMessage("Configurações do registro foram atualizadas com sucesso!", false, forceSilent);
            }
            catch (Exception ex)
            {
                LogMessage($"Erro ao atualizar o registro: {ex.Message}", true, forceSilent);
            }

            // Atualiza a Política de Grupo silenciosamente
            RunProcessAsAdmin("cmd.exe", "/c gpupdate /force /quiet >nul 2>&1", forceSilent);

            LogMessage("Windows Update foi desabilitado com sucesso!", false, forceSilent);

            if (!silentMode && !forceSilent && mainForm != null)
            {
                mainForm.InvokeIfRequired(() => {
                    MessageBox.Show("Windows Update foi desabilitado com sucesso!", "Sucesso",
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                });
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Erro durante o processo de desabilitar o Windows Update: {ex.Message}", true, forceSilent);
            throw;
        }
    }

    public void EnableWindowsUpdate(bool forceSilent = false)
    {
        LogMessage("Iniciando processo de habilitação do Windows Update...", false, forceSilent);

        try
        {
            // Checa se o serviço BITS está em execução
            using (ServiceController bitsService = new ServiceController("BITS"))
            {
                LogMessage("Configurando o serviço BITS...", false, forceSilent);
                RunProcessAsAdmin("cmd.exe", "/c sc config bits start= auto", forceSilent);

                if (bitsService.Status != ServiceControllerStatus.Running)
                {
                    bitsService.Start();
                    bitsService.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                }
                else
                {
                    LogMessage("Serviço BITS já está em execução!", false, forceSilent);
                }
            }

            // Habilita o serviço do Windows Update
            LogMessage("Habilitando o Serviço do Windows Update...", false, forceSilent);
            RunProcessAsAdmin("cmd.exe", "/c sc config wuauserv start= auto", forceSilent);

            using (ServiceController wuService = new ServiceController("wuauserv"))
            {
                if (wuService.Status != ServiceControllerStatus.Running)
                {
                    wuService.Start();
                    wuService.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
                    LogMessage("Serviço Windows Update iniciado com sucesso.", false, forceSilent);
                }
                else
                {
                    LogMessage("Serviço Windows Update já está em execução.", false, forceSilent);
                }
            }

            // Atualiza as configurações do registro
            LogMessage("Atualizando as configurações do registro...", false, forceSilent);
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", true))
                {
                    if (key != null)
                    {
                        key.DeleteValue("NoAutoUpdate", false);
                        key.DeleteValue("AUOptions", false);
                    }
                }
                LogMessage("As configurações do registro foram atualizadas com sucesso!", false, forceSilent);
            }
            catch (Exception ex)
            {
                LogMessage($"Erro ao atualizar o registro: {ex.Message}", true, forceSilent);
            }

            // Atualiza a política de grupo de forma silenciosa
            RunProcessAsAdmin("cmd.exe", "/c gpupdate /force /quiet >nul 2>&1", forceSilent);

            LogMessage("Windows Update foi habilitado com sucesso!", false, forceSilent);

            if (!silentMode && !forceSilent && mainForm != null)
            {
                mainForm.InvokeIfRequired(() => {
                    MessageBox.Show("Windows Update foi habilitado com sucesso!", "Sucesso",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Erro durante o processo de habilitação: {ex.Message}", true, forceSilent);
            throw;
        }
    }

    private void RunProcessAsAdmin(string command, string arguments, bool forceSilent = false)
    {
        using (Process process = new Process())
        {
            process.StartInfo.FileName = command;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            process.StartInfo.StandardOutputEncoding = System.Text.Encoding.GetEncoding(850); // Codificação IBM850 ou CP850
            process.StartInfo.StandardErrorEncoding = System.Text.Encoding.GetEncoding(850);

            LogMessage($"Executando comando: {command} {arguments}", false, forceSilent);

            process.OutputDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    LogMessage($"Saída: {e.Data}", false, forceSilent);
            };

            process.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                    LogMessage($"Erro: {e.Data}", true, forceSilent);
            };

            try
            {
                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();
                process.WaitForExit();

                // Apenas mostra os logs dos exit codes sem sucesso para condições específicas
                if (process.ExitCode != 0 &&
                    !arguments.Contains("gpupdate") &&
                    process.ExitCode != 1056) // 1056 é "serviço já está rodando"
                {
                    LogMessage($"Comando falhou com o código de saída: {process.ExitCode}", true, forceSilent);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Falha ao executar comando: {ex.Message}", true, forceSilent);
            }
        }
    }
}

public class MainForm : Form
{
    private Button toggleButton;
    private RichTextBox logBox;
    private bool isUpdateEnabled = true;
    private MenuStrip menuStrip;
    private WindowsUpdateManager updateManager;

    public MainForm()
    {
        // Cria o update manager com referência para este formulário
        updateManager = new WindowsUpdateManager(this);

        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        this.Icon = Icon.ExtractAssociatedIcon(exePath);

        // Inicializa o formulário
        this.Text = "Gerenciador do Windows Update";
        this.Size = new System.Drawing.Size(600, 500);

        // Cria e configura o MenuStrip
        menuStrip = new MenuStrip();
        ToolStripMenuItem aboutMenuItem = new ToolStripMenuItem("Sobre");
        aboutMenuItem.Click += AboutMenuItem_Click;
        menuStrip.Items.Add(aboutMenuItem);
        this.MainMenuStrip = menuStrip;
        this.Controls.Add(menuStrip);

        // Cria e configura o botão
        toggleButton = new Button();
        toggleButton.Text = "Desabilita o Windows Update";
        toggleButton.Size = new System.Drawing.Size(200, 30);
        toggleButton.Location = new System.Drawing.Point(200, 40);
        toggleButton.Click += ToggleButton_Click;

        // Cria a configura a caixa de logs
        logBox = new RichTextBox();
        logBox.Size = new System.Drawing.Size(560, 300);
        logBox.Location = new System.Drawing.Point(20, 80);
        logBox.ReadOnly = true;
        logBox.BackColor = Color.Black;
        logBox.ForeColor = Color.LightGreen;
        logBox.Font = new Font("Consolas", 10);

        // Adiciona controles ao formulário
        this.Controls.Add(toggleButton);
        this.Controls.Add(logBox);

        // Log do status inicial
        LogMessage("Aplicação inicializada. Checando o status do Windows Update atual...");
        CheckCurrentStatus();
    }

    // Metodo publico para o WindowsUpdateManager utilizar
    public void LogToRichTextBox(string message, bool isError = false)
    {
        LogMessage(message, isError);
    }

    private void AboutMenuItem_Click(object sender, EventArgs e)
    {
        string aboutText = "Gerenciador do Windows Update\n" +
                            "Versão 1.0\n" +
                            "© 2025 - viniciusvalverde.seg.br\n\n" +
                             "Este aplicativo permite habilitar ou desabilitar as atualizações " +
                             "automáticas do Windows de forma fácil e segura.";
        MessageBox.Show(aboutText,
                        "Sobre o Gerenciador do Windows Update",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
    }

    private void LogMessage(string message, bool isError = false)
    {
        string timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
        string logMessage = $"[{timestamp}] {message}\n";

        if (logBox.InvokeRequired)
        {
            logBox.Invoke(new Action(() => {
                logBox.SelectionStart = logBox.TextLength;
                logBox.SelectionLength = 0;
                logBox.SelectionColor = isError ? Color.Red : Color.LightGreen;
                logBox.AppendText(logMessage);
                logBox.ScrollToCaret();
            }));
        }
        else
        {
            logBox.SelectionStart = logBox.TextLength;
            logBox.SelectionLength = 0;
            logBox.SelectionColor = isError ? Color.Red : Color.LightGreen;
            logBox.AppendText(logMessage);
            logBox.ScrollToCaret();
        }
    }

    private void CheckCurrentStatus()
    {
        try
        {
            using (ServiceController service = new ServiceController("wuauserv"))
            {
                // Mostra nos logs o atual status do serviço
                LogMessage($"Status do Serviço Windows Update: {service.Status}");

                // Checa a chave do registro das configurações do Windows Update
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"))
                {
                    if (key != null)
                    {
                        var noAutoUpdate = key.GetValue("NoAutoUpdate");

                        // Mostra nos logs o status da configuração do registro
                        LogMessage($"Configuração do AutoUpdate no registro: {(noAutoUpdate != null ? noAutoUpdate.ToString() : "Não Configurado")}");

                        // Checa se o Windows Update está desabilitado (NoAutoUpdate = 1)
                        if (noAutoUpdate != null && noAutoUpdate.ToString() == "1")
                        {
                            // Atualiza o status da interface gráfica
                            isUpdateEnabled = false;
                            this.InvokeIfRequired(() =>
                            {
                                toggleButton.Text = "Habilitar o Windows Update";
                            });

                            // Mostra os logs do status de desabilitado
                            LogMessage("Status atual: Windows Update DESABILITADO", false);

                            // Mostra a informação no Message Box
                            MessageBox.Show(
                                "Windows Update já encontra-se desabilitado!",
                                "Status do Windows Update",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                            );
                        }
                        else
                        {
                            // Atualiza o estado da interface gráfica
                            isUpdateEnabled = true;
                            this.InvokeIfRequired(() =>
                            {
                                toggleButton.Text = "Desabilitar o Windows Update";
                            });

                            // Mostra o log do status de habilitado
                            LogMessage("Status atual: Windows Update HABILITADO", false);
                        }
                    }
                    else
                    {
                        // Mostra o log quando a chave de registro não foi encontrada
                        LogMessage("Chave do registro não encontrada. Windows Update está no estado padrão.", false);
                        isUpdateEnabled = true;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Mostra o log se algum erro ocorrer durante a checagem
            LogMessage($"Erro ao verificar o status do Windows Update: {ex.Message}", true);
            MessageBox.Show(
                $"Erro ao verificar o status do Windows Update: {ex.Message}",
                "Erro",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private async void ToggleButton_Click(object sender, EventArgs e)
    {
        toggleButton.Enabled = false;
        try
        {
            if (isUpdateEnabled)
            {
                await Task.Run(() => updateManager.DisableWindowsUpdate());
                toggleButton.Text = "Habilitar o Windows Update";
                isUpdateEnabled = false;
            }
            else
            {
                await Task.Run(() => updateManager.EnableWindowsUpdate());
                toggleButton.Text = "Desabilitar o Windows Update";
                isUpdateEnabled = true;
            }
        }
        catch (Exception ex)
        {
            LogMessage($"Erro: {ex.Message}", true);
            MessageBox.Show($"Erro: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            toggleButton.Enabled = true;
        }
    }

    [STAThread]
    static void Main(string[] args)
    {
        bool isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent())
            .IsInRole(WindowsBuiltInRole.Administrator);

        bool silentMode = args.Contains("-silent") || args.Contains("-silently");
        string action = null;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i].ToLower() == "-action" && i + 1 < args.Length)
            {
                action = args[i + 1].ToLower();
                break;
            }
        }

        if (!isAdmin)
        {
            // Reinicia a aplicação com privilégios administrativos
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = true;
            startInfo.WorkingDirectory = Environment.CurrentDirectory;
            startInfo.FileName = Application.ExecutablePath;
            startInfo.Verb = "runas";

            // Passa ao longo qualquer argumento na linha de comando 
            if (args.Length > 0)
            {
                startInfo.Arguments = string.Join(" ", args);
            }

            try
            {
                Process.Start(startInfo);
            }
            catch (Exception)
            {
                if (!silentMode)
                {
                    MessageBox.Show("Esta aplicação requer privilégios de administrador para funcionar.",
                               "Direitos de Administrador Requeridos",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Warning);
                }
                else
                {
                    File.AppendAllText("UpdateManager.log",
                        $"[{DateTime.Now}] ERRO: Falha ao conseguir privilegios administrativos\n");
                }
                return;
            }
            return;
        }

        // Lida com o modo silencioso com ação específica
        if (silentMode && action != null)
        {
            try
            {
                // Cria o WindowsUpdateManager com nenhuma referência ao formulário (silent mode)
                WindowsUpdateManager manager = new WindowsUpdateManager();

                if (action == "enable")
                {
                    manager.EnableWindowsUpdate(true);
                    Environment.Exit(0);
                }
                else if (action == "disable")
                {
                    manager.DisableWindowsUpdate(true);
                    Environment.Exit(0);
                }
                else
                {
                    File.AppendAllText("UpdateManager.log",
                        $"[{DateTime.Now}] ERRO: Acao invalida '{action}'. Use 'enable' ou 'disable'.\n");
                    Environment.Exit(1);
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText("UpdateManager.log",
                    $"[{DateTime.Now}] ERRO: {ex.Message}\n");
                Environment.Exit(1);
            }
        }
        else
        {
            Application.EnableVisualStyles();
            Application.Run(new MainForm());
        }
    }
}