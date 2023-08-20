using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Net;
using System.Management;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static WinFormsApp1.MainForm;
using System;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinFormsApp1
{
    public class WelcomeForm : Form
    {
        private Label messageLabel;
        private Button confirmButton;

        public WelcomeForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

            messageLabel = new Label();
            messageLabel.Location = new System.Drawing.Point(50, 20);
            messageLabel.AutoSize = true;
            messageLabel.Text = "注意：本程序仅限RETE七日杀客户端使用！";

            confirmButton = new Button();
            confirmButton.Location = new System.Drawing.Point(125, 60);
            confirmButton.Text = "继续";
            confirmButton.Click += ConfirmButton_Click;

            Controls.Add(messageLabel);
            Controls.Add(confirmButton);

            Text = "注意！";
            Size = new System.Drawing.Size(350, 150);
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // 设置窗口边框为固定的样式
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK; // 设置窗体的 DialogResult 为 OK
            Close(); // 关闭当前窗体
               }
    }
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            // 创建按钮
            Button button1 = new Button();
            button1.Text = "切换为离线模式";
            button1.Location = new System.Drawing.Point(50, 50);
            button1.Click += Button1_Click; // 按钮1的点击事件
            button1.AutoSize = true;

            Button button2 = new Button();
            button2.Text = "切换为在线模式";
            button2.Location = new System.Drawing.Point(50, 100);
            button2.Click += Button2_Click; // 按钮2的点击事件
            button2.AutoSize = true;

            Button button3 = new Button();
            button3.Text = "更多";
            button3.Location = new System.Drawing.Point(50, 150);
            button3.Click += Button3_Click; // 按钮3的点击事件
            button3.AutoSize = true;

            Label versionLabel = new Label();
            versionLabel.Location = new System.Drawing.Point(190, 340);
            versionLabel.AutoSize = true;

            Button updateButton = new Button();
            updateButton.Location = new Point(50, 200);
            updateButton.Text = "手动更新客户端";
            updateButton.Click += UpdateButton_Click;
            updateButton.AutoSize = true;

            // 获取版本号的逻辑
            string version = "1.0.2"; // 替换为获取版本号的实际逻辑
            versionLabel.Text = "版本号: V" + version;

            Button updateLogButton = new Button();
            updateLogButton.Text = "更新日志";
            updateLogButton.Location = new Point(110, 335
                );
            updateLogButton.AutoSize = true;
            updateLogButton.Click += UpdateLogButton_Click;


            // 添加按钮到窗体
            Controls.Add(button1);
            Controls.Add(button2);
            Controls.Add(button3);
            Controls.Add(versionLabel);
            Controls.Add(updateButton);
            Controls.Add(updateLogButton);

            // 设置窗体属性
            Text = "RETE客户端修改";
            Size = new System.Drawing.Size(300, 400);
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // 设置窗口边框为固定的样式
        }
        private void UpdateButton_Click(object sender, EventArgs e)
        {
            string javaPath = "java"; // Java 可执行文件的路径

            ProcessStartInfo processInfo = new ProcessStartInfo(javaPath, $"-jar .minecraft/updateLoader.jar");
            Process.Start(processInfo);
        }
        private void UpdateLogButton_Click(object sender, EventArgs e)
        {
            // 创建更新日志窗口并显示
            UpdateLogForm updateLogForm = new UpdateLogForm();
            updateLogForm.ShowDialog();
        }
        public class UpdateLogForm : Form
        {
            public UpdateLogForm()
            {
                InitializeComponents();
            }

            private void InitializeComponents()
            {
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
                // 在此处添加更新日志窗口的控件和布局
                // 例如，可以使用 Label 控件显示更新日志的文本
                Label updateLogLabel = new Label();
                updateLogLabel.Text = "1.更新自动更新文件名\n2.修复已知问题";
                updateLogLabel.Location = new Point(20, 20);
                updateLogLabel.AutoSize = true;
                Controls.Add(updateLogLabel);

                // 设置窗口的标题、大小等属性
                Text = "更新日志";
                Size = new Size(400, 300);
            }
        }


        public class DowFrom : Form
        {
            private Label messageLabel;
            private Button confirmButton;

            public DowFrom()
            {
                InitializeComponents();
            }

            private void InitializeComponents()
            {
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

                messageLabel = new Label();
                messageLabel.Location = new System.Drawing.Point(50, 20);
                messageLabel.AutoSize = true;
                messageLabel.Text = "文件下载完成！";

                confirmButton = new Button();
                confirmButton.Location = new System.Drawing.Point(125, 60);
                confirmButton.Text = "安装";
                confirmButton.Click += ConfirmButton_Click;

                Controls.Add(messageLabel);
                Controls.Add(confirmButton);

                Text = "安装Java8";
                Size = new System.Drawing.Size(350, 150);
                this.FormBorderStyle = FormBorderStyle.FixedSingle; // 设置窗口边框为固定的样式
            }

            private void ConfirmButton_Click(object sender, EventArgs e)
            {
                DialogResult = DialogResult.OK; // 设置窗体的 DialogResult 为 OK
                Close(); // 关闭当前窗体
            }
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string configFilePath = Path.Combine(".minecraft", "versions", "RETE1.12.2", "hmclversion.cfg");
                string statusFilePath = Path.Combine(".minecraft", "resources", "rete", "textures", "status.txt");
                string setupFilePath = Path.Combine(".minecraft", "versions", "RETE1.12.2", "PCL", "Setup.ini");

                string statusContent = "离线";

                // 修改 hmclversion.cfg 文件
                JObject configData = JObject.Parse(File.ReadAllText(configFilePath));
                configData["javaArgs"] = "";
                File.WriteAllText(configFilePath, configData.ToString());

                // 修改 status.txt 文件
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                File.WriteAllText(statusFilePath, statusContent, Encoding.GetEncoding("gb2312"));

                // 修改 Setup.ini 文件
                string setupContent = File.ReadAllText(setupFilePath);
                setupContent = Regex.Replace(setupContent, @"(VersionServerLogin:)(\d+)", "VersionServerLogin:0");
                int versionAdvanceJvmIndex = setupContent.IndexOf("VersionAdvanceJvm:");

                if (versionAdvanceJvmIndex != -1)
                {
                    // 定位到 VersionAdvanceJvm 行的末尾
                    int lineEndIndex = setupContent.IndexOf('\n', versionAdvanceJvmIndex);

                    if (lineEndIndex != -1)
                    {
                        // 获取 VersionAdvanceJvm 行的内容
                        string versionAdvanceJvmLine = setupContent.Substring(versionAdvanceJvmIndex, lineEndIndex - versionAdvanceJvmIndex).Trim();

                        // 截取冒号后的内容
                        int colonIndex = versionAdvanceJvmLine.IndexOf(':');
                        string versionAdvanceJvmValue = versionAdvanceJvmLine.Substring(colonIndex + 1).Trim();

                        // 替换内容为所需的值
                        string modifiedVersionAdvanceJvmLine = versionAdvanceJvmLine.Substring(0, colonIndex + 1) + " ";

                        // 替换原始内容为修改后的内容
                        setupContent = setupContent.Replace(versionAdvanceJvmLine, modifiedVersionAdvanceJvmLine);
                    }
                }
                File.WriteAllText(setupFilePath, setupContent);

                MessageBox.Show("成功切换为离线模式");
            }
            catch (Exception ex)
            {
                MessageBox.Show("切换为离线模式时发生错误：" + ex.Message);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                string configFilePath = Path.Combine(".minecraft", "versions", "RETE1.12.2", "hmclversion.cfg");
                string statusFilePath = Path.Combine(".minecraft", "resources", "rete", "textures", "status.txt");
                string setupFilePath = Path.Combine(".minecraft", "versions", "RETE1.12.2", "PCL", "Setup.ini");

                string statusContent = "在线";

                // 修改 hmclversion.cfg 文件
                JObject configData = JObject.Parse(File.ReadAllText(configFilePath));
                configData["javaArgs"] = "-javaagent:updateLoader.jar ";
                File.WriteAllText(configFilePath, configData.ToString());

                // 修改 status.txt 文件
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                File.WriteAllText(statusFilePath, statusContent, Encoding.GetEncoding("gb2312"));

                // 修改 Setup.ini 文件
                string setupContent = File.ReadAllText(setupFilePath);
                setupContent = Regex.Replace(setupContent, @"(VersionServerLogin:)(\d+)", "VersionServerLogin:4");
                int versionAdvanceJvmIndex = setupContent.IndexOf("VersionAdvanceJvm:");

                if (versionAdvanceJvmIndex != -1)
                {
                    // 定位到 VersionAdvanceJvm 行的末尾
                    int lineEndIndex = setupContent.IndexOf('\n', versionAdvanceJvmIndex);

                    if (lineEndIndex != -1)
                    {
                        // 获取 VersionAdvanceJvm 行的内容
                        string versionAdvanceJvmLine = setupContent.Substring(versionAdvanceJvmIndex, lineEndIndex - versionAdvanceJvmIndex).Trim();

                        // 截取冒号后的内容
                        int colonIndex = versionAdvanceJvmLine.IndexOf(':');
                        string versionAdvanceJvmValue = versionAdvanceJvmLine.Substring(colonIndex + 1).Trim();

                        // 替换内容为所需的值
                        string modifiedVersionAdvanceJvmLine = versionAdvanceJvmLine.Substring(0, colonIndex + 1) + "-javaagent:updateLoader.jar ";

                        // 替换原始内容为修改后的内容
                        setupContent = setupContent.Replace(versionAdvanceJvmLine, modifiedVersionAdvanceJvmLine);
                    }
                }
                File.WriteAllText(setupFilePath, setupContent);

                MessageBox.Show("成功切换为在线模式");
            }
            catch (Exception ex)
            {
                MessageBox.Show("切换为在线模式时发生错误：" + ex.Message);
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            SecondForm secondForm = new SecondForm();
            secondForm.ShowDialog();
        }
    }
    public class SecondForm : Form
    {
        private Button button4;
        private Button button5;

        public SecondForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

            // 创建按钮
            button4 = new Button();
            button4.Text = "检查网络";
            button4.Location = new System.Drawing.Point(50, 50);
            button4.Click += Button4_Click; // 按钮4的点击事件

            button5 = new Button();
            button5.Text = "下载Java8";
            button5.Location = new System.Drawing.Point(50, 100);
            button5.Click += Button5_Click; // 按钮5的点击事件

            // 添加按钮到窗体
            Controls.Add(button4);
            Controls.Add(button5);

            // 设置窗体属性
            Text = "更多";
            Size = new System.Drawing.Size(300, 300);
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // 设置窗口边框为固定的样式
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            var networkCheckForm = new NetworkCheckForm();
            networkCheckForm.ShowDialog();
        }

        private async void Button5_Click(object sender, EventArgs e)
        {
            var progressForm = new ProgressForm();
            progressForm.Show();

            var source = "官方源（Github）";

            using (var httpClient = new HttpClient())
            {
                var response = await httpClient.GetAsync("https://api.bell-sw.com/v1/liberica/releases?version-feature=8&version-modifier=latest&bitness=64&fx=false&os=windows&arch=x86&installation-type=installer&package-type=msi&bundle-type=jre");

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    var downloadUrl = ExtractDownloadUrlFromJson(json);

                    var success = false;

                    // 尝试使用原始的 downloadUrl 下载文件
                    success = await DownloadFile(downloadUrl, progressForm, source);

                    // 如果下载失败，则替换为 kgithub.com 后再次尝试下载
                    if (!success)
                    {
                        source = "加速源（kgithub）";
                        var kgithubUrl = ReplaceDomain(downloadUrl, "github.com", "kgithub.com");
                        success = await DownloadFile(kgithubUrl, progressForm, source);
                    }

                    // 如果还是下载失败，则替换为 reteserver.ml/github.com 后再次尝试下载
                    if (!success)
                    {
                        source = "加速源（RETE）";
                        var reteserverUrl = ReplaceDomain(downloadUrl, "kgithub.com", "reteserver.ml/github.com");
                        success = await DownloadFile(reteserverUrl, progressForm, source);
                    }

                    if (success)
                    {
                        DowFrom dowForm = new DowFrom();
                        if (dowForm.ShowDialog() == DialogResult.OK)
                        {
                            var savePath = Path.Combine(Path.GetTempPath(), "downloaded_file.msi");
                            // 使用 msiexec 打开 MSI 文件
                            var processInfo = new ProcessStartInfo("msiexec", $"/i \"{savePath}\"");
                            Process.Start(processInfo);
                        }
                    }
                    else
                    {
                        MessageBox.Show("无法下载文件。");
                    }
                }
                else
                {
                    MessageBox.Show("请求API失败！");
                }
            }

            progressForm.Close();
        }

        private async Task<bool> DownloadFile(string downloadUrl, ProgressForm progressForm, string source)
        {
            using (var downloadClient = new HttpClient())
            {
                downloadClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");

                try
                {
                    var downloadStream = await downloadClient.GetStreamAsync(downloadUrl);
                    var totalBytes = await GetDownloadSize(downloadUrl);

                    var savePath = Path.Combine(Path.GetTempPath(), "downloaded_file.msi");

                    using (var fileStream = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                    {
                        var buffer = new byte[8192];
                        var bytesRead = 0;
                        var totalBytesRead = 0;

                        while ((bytesRead = await downloadStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            totalBytesRead += bytesRead;

                            progressForm.UpdateProgress((int)totalBytesRead, (int)totalBytes, source);
                        }
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    return false;
                }
            }
        }

        private string ReplaceDomain(string downloadUrl, string rurl, string newDomain)
        {
            if (downloadUrl.Contains(rurl))
            {
                downloadUrl = downloadUrl.Replace(rurl, newDomain);
            }
            return downloadUrl;
        }


        private string ExtractDownloadUrlFromJson(string json)
        {
            // 解析json，提取downloadUrl
            var result = Newtonsoft.Json.JsonConvert.DeserializeObject<System.Collections.Generic.List<ApiData>>(json);
            if (result != null && result.Count > 0)
            {
                return result[0].downloadUrl;
            }
            return null;
        }

        private async Task<long> GetDownloadSize(string url)
        {
            using (var httpClient = new HttpClient())
            {
                var headRequest = new HttpRequestMessage(HttpMethod.Head, url);
                var response = await httpClient.SendAsync(headRequest, HttpCompletionOption.ResponseHeadersRead);
                if (response.IsSuccessStatusCode && response.Content.Headers.ContentLength.HasValue)
                {
                    return response.Content.Headers.ContentLength.Value;
                }
            }
            return -1;
        }

        private class ApiData
        {
            public string downloadUrl { get; set; }
        }
        public class ProgressForm : Form
        {
            private ProgressBar progressBar;
            private Label progressLabel;
            private Label sourceLabel;

            private Stopwatch stopwatch;
            private long lastBytesReceived;

            public ProgressForm()
            {
                InitializeComponents();
            }

            private void InitializeComponents()
            {
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

                progressBar = new ProgressBar();
                progressBar.Location = new System.Drawing.Point(50, 50);
                progressBar.Width = 200;

                progressLabel = new Label();
                progressLabel.Location = new System.Drawing.Point(50, 20);
                progressLabel.AutoSize = true;
                progressLabel.AutoEllipsis = true;

                sourceLabel = new Label();
                sourceLabel.Location = new System.Drawing.Point(50, 80);
                sourceLabel.AutoSize = true;  // 自动调整大小以适应文本内容
                sourceLabel.AutoEllipsis = true;  // 超出显示区域时使用省略号

                Controls.Add(progressBar);
                Controls.Add(progressLabel);
                Controls.Add(sourceLabel);

                Text = "进度";
                Size = new System.Drawing.Size(300, 200);
                this.FormBorderStyle = FormBorderStyle.FixedSingle; // 设置窗口边框为固定的样式

                stopwatch = new Stopwatch();
                lastBytesReceived = 0;
            }

            public void UpdateProgress(int current, int total, string source)
            {
                progressBar.Maximum = total;
                progressBar.Value = current;

                double receivedMB = current / (1024.0 * 1024.0);
                double totalMB = total / (1024.0 * 1024.0);

                double speedMBps = (current - lastBytesReceived) / (1024.0 * 1024.0 * stopwatch.Elapsed.TotalSeconds);
                lastBytesReceived = current;

                string progressText = $"{receivedMB:F2}MB/{totalMB:F2}MB";
                progressLabel.Text = progressText;

                sourceLabel.Text = $"当前源：{source}";

                stopwatch.Restart();
            }
        }
        public class NetworkCheckForm : Form
        {
            private Label networkStatusLabel;
            private Label baiduStatusLabel;
            private Label bingStatusLabel;
            private Label skinStatusLabel;
            private Label bbsStatusLabel;
            private Label pclStatusLabel;

            private System.Windows.Forms.Timer timer;

            public NetworkCheckForm()
            {

                InitializeComponents();

                // 创建定时器并设置时间间隔为1秒
                timer = new System.Windows.Forms.Timer();

                // 设置定时器的Tick事件处理程序
                timer.Tick += Timer_Tick;

                // 启动定时器
                timer.Start();
            }

            private void InitializeComponents()
            {
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));

                networkStatusLabel = new Label();
                networkStatusLabel.Location = new System.Drawing.Point(50, 20);
                networkStatusLabel.AutoSize = true;
                networkStatusLabel.Text = "网络连接：未检测";

                baiduStatusLabel = new Label();
                baiduStatusLabel.Location = new System.Drawing.Point(50, 50);
                baiduStatusLabel.AutoSize = true;
                baiduStatusLabel.Text = "百度连接：未检测";

                bingStatusLabel = new Label();
                bingStatusLabel.Location = new System.Drawing.Point(50, 80);
                bingStatusLabel.AutoSize = true;
                bingStatusLabel.Text = "Bing连接：未检测";

                skinStatusLabel = new Label();
                skinStatusLabel.Location = new System.Drawing.Point(50, 110);
                skinStatusLabel.AutoSize = true;
                skinStatusLabel.Text = "皮肤站连接：未检测";

                bbsStatusLabel = new Label();
                bbsStatusLabel.Location = new System.Drawing.Point(50, 140);
                bbsStatusLabel.AutoSize = true;
                bbsStatusLabel.Text = "论坛连接：未检测";

                pclStatusLabel = new Label();
                pclStatusLabel.Location = new System.Drawing.Point(50, 170);
                pclStatusLabel.AutoSize = true;
                pclStatusLabel.Text = "PCL主页连接：未检测";

                Controls.Add(networkStatusLabel);
                Controls.Add(baiduStatusLabel);
                Controls.Add(bingStatusLabel);
                Controls.Add(skinStatusLabel);
                Controls.Add(bbsStatusLabel);
                Controls.Add(pclStatusLabel);

                Text = "网络连接检查";
                Size = new System.Drawing.Size(400, 300);
                this.FormBorderStyle = FormBorderStyle.FixedSingle; // 设置窗口边框为固定的样式
            }

            private async void Timer_Tick(object sender, EventArgs e)
            {
                timer.Stop(); // 停止定时器

                bool isConnected = NetworkInterface.GetIsNetworkAvailable();
                networkStatusLabel.Text = isConnected ? "网络连接：正常" : "网络连接：失败";

                bool baiduConnection = false;
                bool bingConnection = false;
                bool skinConnection = false;
                bool bbsConnection = false;
                bool pclConnection = false;
                bool Dns = false;

                if (isConnected)
                {
                    baiduConnection = await CheckWebsiteConnection("https://www.baidu.com", baiduStatusLabel, "百度连接");
                    bingConnection = await CheckWebsiteConnection("https://www.bing.com", bingStatusLabel, "Bing连接");
                    skinConnection = await CheckWebsiteConnection("https://skin.mcrete.top", skinStatusLabel, "皮肤站连接");
                    bbsConnection = await CheckWebsiteConnection("https://bbs.mcrete.top", bbsStatusLabel, "论坛连接");
                    pclConnection = await CheckWebsiteConnection("https://reteserver.ml", pclStatusLabel, "PCL主页连接");
                }
                else
                {
                    baiduStatusLabel.Text = "百度连接：未检测";
                    bingStatusLabel.Text = "Bing连接：未检测";
                    skinStatusLabel.Text = "皮肤站连接：未检测";
                    bbsStatusLabel.Text = "论坛连接：未检测";
                    pclStatusLabel.Text = "PCL主页连接：未检测";
                }

                if ((!baiduConnection || !bingConnection || !skinConnection || !bbsConnection || !pclConnection) && !Dns)
                {
                    Dns = true;
                    ShowDnsChangeConfirmation();
                    timer.Stop(); // 停止定时器
                }
                else
                {
                    timer.Start(); // 重新启动定时器
                }
            }

            private async Task<bool> CheckWebsiteConnection(string url, Label statusLabel, string labelName)
            {
                try
                {
                    using (var client = new WebClient())
                    {
                        var downloadTask = client.DownloadStringTaskAsync(url);
                        var timeoutTask = Task.Delay(TimeSpan.FromSeconds(10));

                        var completedTask = await Task.WhenAny(downloadTask, timeoutTask);

                        if (completedTask == downloadTask)
                        {
                            // 下载任务完成
                            statusLabel.Text = $"{labelName}：正常";
                            return true;
                        }
                        else
                        {
                            // 超时任务完成
                            statusLabel.Text = $"{labelName}：超时";
                            return false;
                        }
                    }
                }
                catch (Exception)
                {
                    statusLabel.Text = $"{labelName}：失败";
                    return false;
                }
            }

            private void ShowDnsChangeConfirmation()
            {
                DialogResult result = MessageBox.Show("网站连接异常，是否更换DNS?", "提示", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    // 执行更换DNS的操作
                    ChangeDnsSettings();
                }
                else
                {
                    // 关闭窗口
                    Close();
                }
            }

            private void ChangeDnsSettings()
            {
                SetDNS(new string[] { "223.5.5.5", "223.6.6.6" });
                MessageBox.Show("DNS已更换为阿里DNS", "提示", MessageBoxButtons.OK);
            }
            public static void SetDNS(string[] dns)
            {
                ManagementClass wmi = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = wmi.GetInstances();
                ManagementBaseObject inPar = null;
                ManagementBaseObject outPar = null;
                foreach (ManagementObject mo in moc)
                {
                    //如果没有启用 IP 设置的网络设备则跳过
                    if (!(bool)mo["IPEnabled"])
                        continue;

                    //设置 DNS 地址
                    if (dns != null)
                    {
                        inPar = mo.GetMethodParameters("SetDNSServerSearchOrder");
                        inPar["DNSServerSearchOrder"] = dns;
                        outPar = mo.InvokeMethod("SetDNSServerSearchOrder", inPar, null);
                    }
                    else if (dns[0] == "back")
                    {
                        mo.InvokeMethod("SetDNSServerSearchOrder", null);
                    }
                }
            }
        }
    }
}