using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BatchOfficeTool
{
    public partial class MainForm : Form
    {
        private List<string> selectedFiles = new List<string>();
        private FindReplaceEngine findReplaceEngine;
        private PrintEngine printEngine;

        public MainForm()
        {
            InitializeComponent();
            findReplaceEngine = new FindReplaceEngine();
            printEngine = new PrintEngine();
        }

        private void InitializeComponent()
        {
            this.Text = "批量Office工具 - 查找替换和打印";
            this.Size = new System.Drawing.Size(900, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = System.Drawing.Color.WhiteSmoke;

            // 主容器
            var mainPanel = new Panel() { Dock = DockStyle.Fill, Padding = new Padding(10) };

            // 标签页
            var tabControl = new TabControl() { Dock = DockStyle.Fill };

            // 标签页1：查找替换
            var tabFindReplace = new TabPage("查找替换");
            tabFindReplace.Controls.Add(CreateFindReplaceTab());

            // 标签页2：批量打印
            var tabPrint = new TabPage("批量打印");
            tabPrint.Controls.Add(CreatePrintTab());

            tabControl.TabPages.Add(tabFindReplace);
            tabControl.TabPages.Add(tabPrint);

            mainPanel.Controls.Add(tabControl);
            this.Controls.Add(mainPanel);
        }

        private Panel CreateFindReplaceTab()
        {
            var panel = new Panel() { Dock = DockStyle.Fill, Padding = new Padding(10), AutoScroll = true };

            var y = 10;

            // 文件选择区域
            var lblFiles = new Label() { Text = "选择文件:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold) };
            panel.Controls.Add(lblFiles);
            y += 30;

            var btnSelectFiles = new Button() { Text = "选择 Word/Excel 文件", Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(150, 35), BackColor = System.Drawing.Color.SkyBlue };
            btnSelectFiles.Click += BtnSelectFiles_Click;
            panel.Controls.Add(btnSelectFiles);

            var btnClearFiles = new Button() { Text = "清空列表", Location = new System.Drawing.Point(170, y), Size = new System.Drawing.Size(100, 35), BackColor = System.Drawing.Color.LightCoral };
            btnClearFiles.Click += BtnClearFiles_Click;
            panel.Controls.Add(btnClearFiles);
            y += 45;

            // 文件列表
            var fileListBox = new ListBox() { Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(850, 120) };
            this.fileListBox = fileListBox;
            panel.Controls.Add(fileListBox);
            y += 130;

            // 查找替换参数
            var lblFind = new Label() { Text = "查找内容:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold) };
            panel.Controls.Add(lblFind);

            var txtFind = new TextBox() { Location = new System.Drawing.Point(100, y), Size = new System.Drawing.Size(300, 25) };
            this.txtFind = txtFind;
            panel.Controls.Add(txtFind);
            y += 35;

            var lblReplace = new Label() { Text = "替换为:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold) };
            panel.Controls.Add(lblReplace);

            var txtReplace = new TextBox() { Location = new System.Drawing.Point(100, y), Size = new System.Drawing.Size(300, 25) };
            this.txtReplace = txtReplace;
            panel.Controls.Add(txtReplace);
            y += 35;

            var chkCaseSensitive = new CheckBox() { Text = "区分大小写", Location = new System.Drawing.Point(10, y), AutoSize = true };
            this.chkCaseSensitive = chkCaseSensitive;
            panel.Controls.Add(chkCaseSensitive);
            y += 30;

            // 执行按钮
            var btnExecute = new Button() { Text = "执行查找替换", Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(150, 40), BackColor = System.Drawing.Color.LimeGreen, Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold) };
            btnExecute.Click += BtnExecute_Click;
            panel.Controls.Add(btnExecute);
            y += 50;

            // 日志显示
            var lblLog = new Label() { Text = "操作日志:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold) };
            panel.Controls.Add(lblLog);
            y += 25;

            var logBox = new RichTextBox() { Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(850, 150), ReadOnly = true };
            this.logBox = logBox;
            panel.Controls.Add(logBox);

            return panel;
        }

        private Panel CreatePrintTab()
        {
            var panel = new Panel() { Dock = DockStyle.Fill, Padding = new Padding(10), AutoScroll = true };

            var y = 10;

            // 文件选择区域
            var lblFiles = new Label() { Text = "选择要打印的文件:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold) };
            panel.Controls.Add(lblFiles);
            y += 30;

            var btnSelectPrintFiles = new Button() { Text = "选择文件", Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(150, 35), BackColor = System.Drawing.Color.SkyBlue };
            btnSelectPrintFiles.Click += BtnSelectPrintFiles_Click;
            panel.Controls.Add(btnSelectPrintFiles);

            var btnClearPrintFiles = new Button() { Text = "清空列表", Location = new System.Drawing.Point(170, y), Size = new System.Drawing.Size(100, 35), BackColor = System.Drawing.Color.LightCoral };
            btnClearPrintFiles.Click += BtnClearPrintFiles_Click;
            panel.Controls.Add(btnClearPrintFiles);
            y += 45;

            // 打印文件列表
            var printFileListBox = new ListBox() { Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(850, 150) };
            this.printFileListBox = printFileListBox;
            panel.Controls.Add(printFileListBox);
            y += 160;

            // 打印设置
            var chkPrintToFile = new CheckBox() { Text = "打印到文件（PDF）", Location = new System.Drawing.Point(10, y), AutoSize = true };
            this.chkPrintToFile = chkPrintToFile;
            panel.Controls.Add(chkPrintToFile);
            y += 30;

            var lblCopies = new Label() { Text = "打印份数:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 9) };
            panel.Controls.Add(lblCopies);

            var numCopies = new NumericUpDown() { Location = new System.Drawing.Point(100, y), Size = new System.Drawing.Size(80, 25), Minimum = 1, Maximum = 100, Value = 1 };
            this.numCopies = numCopies;
            panel.Controls.Add(numCopies);
            y += 35;

            // 执行打印
            var btnPrint = new Button() { Text = "开始打印", Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(150, 40), BackColor = System.Drawing.Color.Orange, Font = new System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold) };
            btnPrint.Click += BtnPrint_Click;
            panel.Controls.Add(btnPrint);
            y += 50;

            // 打印日志
            var lblPrintLog = new Label() { Text = "打印日志:", Location = new System.Drawing.Point(10, y), Font = new System.Drawing.Font("Arial", 9, System.Drawing.FontStyle.Bold) };
            panel.Controls.Add(lblPrintLog);
            y += 25;

            var printLogBox = new RichTextBox() { Location = new System.Drawing.Point(10, y), Size = new System.Drawing.Size(850, 150), ReadOnly = true };
            this.printLogBox = printLogBox;
            panel.Controls.Add(printLogBox);

            return panel;
        }

        // 查找替换标签页控件
        private ListBox fileListBox;
        private TextBox txtFind;
        private TextBox txtReplace;
        private CheckBox chkCaseSensitive;
        private RichTextBox logBox;

        // 打印标签页控件
        private ListBox printFileListBox;
        private CheckBox chkPrintToFile;
        private NumericUpDown numCopies;
        private RichTextBox printLogBox;

        // 事件处理
        private void BtnSelectFiles_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Office 文件 (*.docx, *.xlsx, *.doc, *.xls)|*.docx;*.xlsx;*.doc;*.xls|Word 文件 (*.docx, *.doc)|*.docx;*.doc|Excel 文件 (*.xlsx, *.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    selectedFiles.Clear();
                    selectedFiles.AddRange(ofd.FileNames);
                    UpdateFileListBox();
                    AddLog($"已选择 {ofd.FileNames.Length} 个文件");
                }
            }
        }

        private void BtnClearFiles_Click(object sender, EventArgs e)
        {
            selectedFiles.Clear();
            UpdateFileListBox();
            AddLog("文件列表已清空");
        }

        private void UpdateFileListBox()
        {
            fileListBox.Items.Clear();
            foreach (var file in selectedFiles)
            {
                fileListBox.Items.Add(Path.GetFileName(file));
            }
        }

        private void BtnExecute_Click(object sender, EventArgs e)
        {
            if (selectedFiles.Count == 0)
            {
                MessageBox.Show("请先选择文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtFind.Text))
            {
                MessageBox.Show("请输入查找内容！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var findText = txtFind.Text;
            var replaceText = txtReplace.Text;
            var caseSensitive = chkCaseSensitive.Checked;

            AddLog($"\n开始执行查找替换操作...");
            AddLog($"查找: '{findText}', 替换为: '{replaceText}', 区分大小写: {caseSensitive}");

            int totalReplaced = 0;

            foreach (var file in selectedFiles)
            {
                try
                {
                    AddLog($"\n处理文件: {Path.GetFileName(file)}");
                    int replaced = 0;

                    if (file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                    {
                        replaced = findReplaceEngine.ReplaceInWord(file, findText, replaceText, caseSensitive);
                    }
                    else if (file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        replaced = findReplaceEngine.ReplaceInExcel(file, findText, replaceText, caseSensitive);
                    }

                    AddLog($"✓ {Path.GetFileName(file)} - 替换 {replaced} 个");
                    totalReplaced += replaced;
                }
                catch (Exception ex)
                {
                    AddLog($"✗ {Path.GetFileName(file)} - 错误: {ex.Message}");
                }
            }

            AddLog($"\n操作完成！总共替换 {totalReplaced} 个内容");
            MessageBox.Show($"操作完成！\n总共替换 {totalReplaced} 个内容", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnSelectPrintFiles_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Filter = "Office 文件 (*.docx, *.xlsx, *.doc, *.xls)|*.docx;*.xlsx;*.doc;*.xls|Word 文件 (*.docx, *.doc)|*.docx;*.doc|Excel 文件 (*.xlsx, *.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    printFileListBox.Items.Clear();
                    foreach (var file in ofd.FileNames)
                    {
                        printFileListBox.Items.Add(file);
                    }
                    AddPrintLog($"已选择 {ofd.FileNames.Length} 个文件");
                }
            }
        }

        private void BtnClearPrintFiles_Click(object sender, EventArgs e)
        {
            printFileListBox.Items.Clear();
            AddPrintLog("打印文件列表已清空");
        }

        private void BtnPrint_Click(object sender, EventArgs e)
        {
            if (printFileListBox.Items.Count == 0)
            {
                MessageBox.Show("请先选择文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AddPrintLog($"\n开始批量打印...");
            int copies = (int)numCopies.Value;
            bool printToPdf = chkPrintToFile.Checked;

            int successCount = 0;
            int failureCount = 0;

            foreach (var item in printFileListBox.Items)
            {
                var file = item.ToString();
                try
                {
                    AddPrintLog($"正在打印: {Path.GetFileName(file)}...");

                    if (file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                    {
                        printEngine.PrintWord(file, copies, printToPdf);
                    }
                    else if (file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        printEngine.PrintExcel(file, copies, printToPdf);
                    }

                    AddPrintLog($"✓ {Path.GetFileName(file)} 打印成功 ({copies} 份)");
                    successCount++;
                }
                catch (Exception ex)
                {
                    AddPrintLog($"✗ {Path.GetFileName(file)} 打印失败: {ex.Message}");
                    failureCount++;
                }
            }

            AddPrintLog($"\n打印完成！成功: {successCount}, 失败: {failureCount}");
            MessageBox.Show($"打印完成！\n成功: {successCount}\n失败: {failureCount}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void AddLog(string message)
        {
            if (logBox.InvokeRequired)
            {
                logBox.Invoke(new Action(() => AddLog(message)));
            }
            else
            {
                logBox.AppendText(message + Environment.NewLine);
                logBox.ScrollToCaret();
            }
        }

        private void AddPrintLog(string message)
        {
            if (printLogBox.InvokeRequired)
            {
                printLogBox.Invoke(new Action(() => AddPrintLog(message)));
            }
            else
            {
                printLogBox.AppendText(message + Environment.NewLine);
                printLogBox.ScrollToCaret();
            }
        }
    }
}
