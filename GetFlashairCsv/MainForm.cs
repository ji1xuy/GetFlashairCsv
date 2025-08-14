using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;
using System.Diagnostics;
using Application = System.Windows.Forms.Application;
using System.Xml;
using System.Net;

//Nuget �p�b�P�[�W�̃C���X�g�[��
//Selenium.WebDriver
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;

//Nuget �p�b�P�[�W�̃C���X�g�[��
//WebDriverManager
using WebDriverManager;
using WebDriverManager.Helpers;
using WebDriverManager.DriverConfigs.Impl;

//Nuget �p�b�P�[�W�̃C���X�g�[��
//DocumentFormat.OpenXml
using DocumentFormat.OpenXml;
using OOXML = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OOXMLS = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.ComponentModel;
//using DocumentFormat.OpenXml.Office.CustomUI;
//using NPOI.OpenXmlFormats.Dml;
//using NPOI.SS.Formula.Functions;
//using System.Reflection.Metadata;
//using AngleSharp.Text;

namespace GetFlashairCsv {
    public partial class MainForm : Form {
        private const string APPNAME = "GetFlashairCsv";
        private const string WINDOW_TITLE = APPNAME + "_20250814";
        private const string INIFILE_FILENAME = @"./" + APPNAME + ".ini"; // "./"�v
        private const string INIFILE_KEY_URL = "url";
        private const string INIFILE_KEY_BROWSER = "browser";
        private const string PROTOCOL = "http://";
        private const string INIFILE_KEY_MAC_ADDR = "macAddr";
        private const string INIFILE_KEY_START_IP_ADDR = "startIpAddr";
        private const string INIFILE_KEY_END_IP_ADDR = "endIpAddr";
        private const string EXCEL_FILENAME = @"whm_30min.xlsx";
        private const string EXCEL_SHEETNAME = "30���f�[�^";
        private const string EXCEL_TABLENAME = "�e�[�u��1";
        private const string EXCEL_TABLESTYLENAME = "TableStyleMedium16";
        private const string ITEM_NOT_SELECTED = "���I��";
        private const string CAPTION_ERROR = "�G���[";
        private const string CAPTION_INFORMATION = "���";
        private const string CAPTION_QUESTION = "�m�F";
        private const int ERROR_RETURN_VALUE = -1;
        //EXCEL_HEADER_IDENTIFIER: CSV�t�@�C���AExcel�t�@�C���̃w�b�_�[(��1���)�̎��ʎq
        private const string EXCEL_HEADER_IDENTIFIER = "yyyy/mm/dd";
        //�J�X�^�������`��������
        //https://learn.microsoft.com/ja-jp/dotnet/standard/base-types/custom-date-and-time-format-strings#H_Specifier
        // M�͌��Am�͕��Ah��12���Ԍ`���̎��ԁAH��24���Ԍ`���̎���
        // 1���͐�s�[���Ȃ��A2���͐�s�[������)
        private const string EXCEL_DATE_FORMAT = "yyyy/MM/dd";
        private const string EXCEL_TIME_FORMAT = "HH:mm";
        private const string EXCEL_STYLE_DATE_FORMATCODE = "yyyy/M/d";
        private const string EXCEL_STYLE_TIME_FORMATCODE = "H:mm;@";
        private const string CLOCK_FORMAT = "yyyy/MM/dd HH:mm:ss";
        private const int OUT_OF_SCREEN = -32000;
        private MainForm mainForm;
        private Flashair flashair;
        private CsvFileList csvFileList;
        private ProgressForm? progressForm = null;
        private System.DateTime lastDateTime;
        private IntPtr? browserHandle = null;
        private FindFlashairForm? findFlashairForm = null;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern uint GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault,
           StringBuilder lpReturnedString, uint nSize, string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool WritePrivateProfileString(string lpAppName,
           string lpKeyName, string lpString, string lpFileName);

        [DllImport("kernel32.dll", EntryPoint = "GetSystemDefaultLCID")]
        public static extern int GetSystemDefaultLCID();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd,
           int hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll")]
        private static extern bool RemoveMenu(IntPtr hMenu, uint uPosition, uint uFlags);

        const uint MF_BYPOSITION = 0x400;
        const uint MF_BYCOMMAND = 0x0;
        const uint SC_CLOSE = 0xF060;

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(
           IntPtr hWnd, out int lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        extern static int SendMessage(IntPtr hwnd, int msg, int wParam, int lParam);

        const int WM_CLOSE = 0x0010;

        [DllImport("iphlpapi.dll", ExactSpelling = true)]
        private static extern int SendARP(int DestIP, int SrcIp, byte[] pMacAddr, ref int PhyAddrLen);

        [DllImport("kernel32.dll", EntryPoint = "FormatMessageA")]
        private static extern int FormatMessage(uint dwFlags, long lpSource, int dwMessageId,
            int dwLanguageId, StringBuilder lpBuffer, uint nSize, long Arguments);

        /*
        GetFlashairCsv.ini�ɂ��ݒ���@
        �e�L�[�̖��O�͒萔�ɂĐݒ�
        FlashAir��CSV�t�@�C���̃��X�g�X�V(GUI�Őݒ�ύX�\)
        �EFlashAir��URL     �L�[: INIFILE_KEY_URL�̒l / �l: http://xxx.xxx.xxx.xxx
        �E�g�p����u���E�U     �L�[: INIFILE_KEY_BROWSER�̒l / �l: Chrome or Edge
        FlashAir�̌���(ini�t�@�C���̒��ڕҏW�̂�)
        �EMAC�A�h���X        �L�[:�@INIFILE_KEY_MAC_ADDR�̒l / �l: xx-xx-xx-xx-xx-xx
        �E�����J�nIP�A�h���X   �L�[:�@INIFILE_KEY_START_IP_ADDR�̒l / �l: nnn.nnn.nnn.nnn
        �E�����I��IP�A�h���X   �L�[:�@INIFILE_KEY_END_IP_ADDR�̒l / �l: nnn.nnn.nnn.nnn
        (�����J�nIP�A�h���X�ƌ����I��IP�A�h���X�͓���Z�O�����g���ł��邱��)
        */

        public MainForm() {
            InitializeComponent();
            mainForm = this;
            Debug.WriteLine("mainForm.Handle: " + mainForm.Handle);
            flashair = new Flashair(this);
            flashair.ReadUrlFromInifile();
            ReadBrowserFromInifile();
            csvFileList = new CsvFileList(this);

            //�^�C�g���o�[�̐ݒ�
            this.Text = WINDOW_TITLE;

            //Excel�t�@�C���̃t���p�X�����x���ɕ\��
            ExcelFileNameLabel.Text =
                Path.Combine(Application.StartupPath, EXCEL_FILENAME);

            WriteExcelButton.BackColor = System.Drawing.Color.Yellow;

            //lastDateTime�̏����l��ݒ�
            var now = System.DateTime.Now;
            //30���ȏ��30���A30��������0���A�b��0�b�A�~���b��0�~���b�ɒ����@2�̕��@
            //new DateTime()�ŐV���ɍ쐬
            lastDateTime = new System.DateTime(
                now.Year, now.Month, now.Day, now.Hour,
                (now.Minute >= 30) ? 30 : 0, 0, 0);
            Debug.WriteLine("lastDateTime: " + lastDateTime);

            timer1_Tick(Type.Missing, EventArgs.Empty); //�����̓_�~�[
        }
        private class Flashair {
            private MainForm _mainForm;

            public Flashair(MainForm mainForm) {
                _mainForm = mainForm;
                _mainForm.FlashairUrlTextBox.Text = PROTOCOL;
            }

            public string? Url {
                get { return _mainForm.FlashairUrlTextBox.Text; }
                private set { _mainForm.FlashairUrlTextBox.Text = value; }
            }

            public string? MacAddr { get; private set; }

            public string? StartIpAddr { get; private set; }

            public string? EndIpAddr { get; private set; }

            public void WriteUrlToInifile() {
                if (WritePrivateProfileString(APPNAME, INIFILE_KEY_URL,
                    _mainForm.FlashairUrlTextBox.Text, INIFILE_FILENAME)) {
                    _mainForm.ShowOKMessageBox("�ۑ����܂���");
                } else {
                    _mainForm.ShowOKMessageBox("�ۑ��ł��܂���ł���", CAPTION_ERROR,
                        MessageBoxIcon.Error);
                }
            }

            public void ReadMacAddrFromInifile() {
                //ini�t�@�C������ݒ��ǂݍ���
                int capacitySize = 256;
                //StringBuilder�N���X
                //������̒ǉ��A�u���A�}�����s���ƁA
                //�I�u�W�F�N�g�̓��e���ύX����邾���ŐV�����I�u�W�F�N�g���쐬���܂���
                var sb = new StringBuilder(capacitySize);
                var stringLength = GetPrivateProfileString(
                    APPNAME, INIFILE_KEY_MAC_ADDR, "", sb, Convert.ToUInt32(sb.Capacity),
                    INIFILE_FILENAME);
                //FLashAir��URL �����x���ɕ\��
                if (stringLength > 0) {
                    MacAddr = sb.ToString();
                } else {
                    MacAddr = "";
                }
            }

            public void ReadStartIpAddrFromInifile() {
                int capacitySize = 256;
                var sb = new StringBuilder(capacitySize);
                var stringLength = GetPrivateProfileString(
                    APPNAME, INIFILE_KEY_START_IP_ADDR, "", sb, Convert.ToUInt32(sb.Capacity),
                    INIFILE_FILENAME);
                if (stringLength > 0) {
                    StartIpAddr = sb.ToString();
                } else {
                    StartIpAddr = "";
                }
            }

            public void ReadEndIpAddrFromInifile() {
                int capacitySize = 256;
                var sb = new StringBuilder(capacitySize);
                var stringLength = GetPrivateProfileString(
                    APPNAME, INIFILE_KEY_END_IP_ADDR, "", sb, Convert.ToUInt32(sb.Capacity),
                    INIFILE_FILENAME);
                if (stringLength > 0) {
                    EndIpAddr = sb.ToString();
                } else {
                    EndIpAddr = "";
                }
            }

            public void ReadUrlFromInifile() {
                int capacitySize = 256;
                var sb = new StringBuilder(capacitySize);
                var stringLength = GetPrivateProfileString(
                    APPNAME, INIFILE_KEY_URL, "", sb, Convert.ToUInt32(sb.Capacity),
                    INIFILE_FILENAME);
                if (stringLength > 0) {
                    _mainForm.FlashairUrlTextBox.Text = sb.ToString();
                }
            }

            public async Task<bool> DownloadCSVFile(ProgressForm progressform) {
                //CSV�t�@�C�������擾
                if (_mainForm.CsvFileNameLabel.Text == ITEM_NOT_SELECTED) {
                    _mainForm.ShowErrorMessageBox("CSV�t�@�C�����I������Ă��܂���");
                    return false;
                }

                //FlashAir��URL��CSV�t�@�C������A��
                var filepath = _mainForm.FlashairUrlTextBox.Text;
                if (filepath.EndsWith("/") == false) {
                    filepath += "/";
                }
                filepath += _mainForm.CsvFileNameLabel.Text;

                //CSV�t�@�C�����_�E�����[�h
                var client = new HttpClient();
                //�^�C���A�E�g���Ԃ̐ݒ�i�f�t�H���g��100�b�j
                client.Timeout = TimeSpan.FromSeconds(30);
                HttpResponseMessage? response = null;
                try {
                    response = await client.GetAsync(filepath);
                } catch (Exception e) when (e is TaskCanceledException || e is HttpRequestException) {
                    //TaskCanceledException: client.Timeout�Őݒ肵���^�C���A�E�g����������
                    //HttpRequestException: �ڑ��ς݂̌Ăяo���悪���̎��Ԃ��߂��Ă��������������Ȃ�����
                    _mainForm.ShowErrorMessageBox(progressform,
                        "�ʐM���Ƀ^�C���A�E�g�������������߃_�E�����[�h�𒆎~���܂���");
                    return false;
                } catch (Exception e) {
                    _mainForm.ShowErrorMessageBox(progressform, e);
                    return false;
                }
                if (response!.StatusCode != System.Net.HttpStatusCode.OK) {
                    _mainForm.ShowErrorMessageBox(progressform,
                        "CSV�t�@�C�����_�E�����[�h�ł��܂���ł���");
                    return false;
                }
                //�ۑ�
                using var stream = await response.Content.ReadAsStreamAsync();
                using var outStream = File.Create(_mainForm.CsvFileNameLabel.Text);
                stream.CopyTo(outStream);
                return true;
            }
        }

        private class CsvFileList {
            private MainForm _mainForm;

            public CsvFileList(MainForm mainForm) {
                _mainForm = mainForm;
                _mainForm.CsvFileListBox.Items.Clear();
                _mainForm.CsvFileNameLabel.Text = ITEM_NOT_SELECTED;
            }

            public string FileName {
                get {
                    return _mainForm.CsvFileNameLabel.Text;
                }
                set {
                    if (value == null) {
                        return;
                    }
                    _mainForm.CsvFileNameLabel.Text = value;
                }
            }

            public async Task<bool> Update(ProgressForm progressForm) {
                return await Task.Run(() => {
                    var list = new List<string>();
                    IWebDriver? driver;
                    if (_mainForm.ChromeRadioButton.Checked) {
                        // Web�h���C�o�[�̃C���X�^���X��
                        ChromeDriverService? chromeService;
                        ChromeOptions chromeOptions = new();
                        chromeOptions.AddArgument("--window-position=" + OUT_OF_SCREEN + "," + OUT_OF_SCREEN);
                        //chromeService = ChromeDriverService.CreateDefaultService();
                        //chromeService = ChromeDriverService.CreateDefaultService(Application.StartupPath);
                        //�h���C�o�̋N���ꏊ�Ɏ����ۑ����ꂽ�ꏊ���w��
                        var driverVersion = new ChromeConfig().GetMatchingBrowserVersion();
                        var driverPath = $"./Chrome/{driverVersion}/X64/";
                        chromeService = ChromeDriverService.CreateDefaultService(driverPath);
                        //chromeService.SuppressInitialDiagnosticInformation = true; //�f�f�o�͗}��
                        chromeService.HideCommandPromptWindow = true; //�R�}���h�v�����v�g��ʔ�\��
                        //chromeOptions.AddArgument("--headless");
                        //Normal: complete(���ׂẴ��\�[�X���_�E�����[�h����̂�҂��܂�)
                        chromeOptions.PageLoadStrategy = PageLoadStrategy.Normal;

                        try {
                            using (driver = new ChromeDriver(chromeService, chromeOptions)) {
                                //Minimize()���ƃu���E�U����u�\������Ă��܂�
                                //driver.Manage().Window.Minimize();
                                //headless���[�h�ł̓n���h�����擾�ł��Ȃ�
                                _mainForm.browserHandle = GetBrowserHandle();
                                progressForm.Invoke((MethodInvoker)(() => {
                                    progressForm.AbortButton.Enabled = true;
                                }));
                                driver.Navigate().GoToUrl(_mainForm.flashair.Url!);
                                ReadOnlyCollection<IWebElement> elms =
                                    driver.FindElements(By.XPath(@"//*[@id='thumbnail']/div"));
                                //30���l�f�[�^��CSV�t�@�C�����̃��X�g���擾
                                // ReadOnlyCollection<IWebElement>�^�̃��X�g����
                                // List<IWebElement>�^�̃��X�g�𐶐�����
                                // ����� ConvertAll() ��List<string>�^�ɕϊ�����
                                list = (new List<IWebElement>(elms)).ConvertAll(elm => elm.Text);
                            }
                        } catch (Exception e) {
                            HandleException(progressForm, e);
                            return false;
                        }
                    }
                    if (_mainForm.EdgeRadioButton.Checked) {
                        EdgeDriverService? edgeService;
                        EdgeOptions edgeOptions = new();
                        //edgeService = EdgeDriverService.CreateDefaultService();
                        //�h���C�o�̋N���ꏊ�Ɏ����ۑ����ꂽ�ꏊ���w��
                        var driverVersion = new EdgeConfig().GetMatchingBrowserVersion();
                        var driverPath = $"./Edge/{driverVersion}/X64/";
                        edgeService = EdgeDriverService.CreateDefaultService(driverPath);

                        edgeService.HideCommandPromptWindow = true;
                        //��� HideCommandPromptWindow = true �ł�
                        //�R�}���h�v�����v�g����u�\������邽��
                        //���̕��@(2�ڂ�Answer)�ɕς��Ė�������ʊO�ɕ\��������悤�ɂ���
                        //https://stackoverflow.com/questions/35818436/hide-silence-chromedriver-window
                        edgeOptions.AddArgument("--window-position=" + OUT_OF_SCREEN + "," + OUT_OF_SCREEN);
                        //Microsoft Edge WebDriver �����Ă��Ȃ������̂����{����
                        //Microsoft Edge WebDriver ������
                        //HideCommandPromptWindow = true �ɖ߂���

                        //edgeOptions.AddArgument("--headless");
                        edgeOptions.PageLoadStrategy = PageLoadStrategy.Normal;
                        //edgeOptions.AddArgument("--user-data-dir=C:\\Users\\aida0\\AppData\\Local\\Microsoft\\Edge\\User Data");
                        //edgeOptions.AddArgument("--profile-directory=Default");
                        try {
                            using (driver = new EdgeDriver(edgeService, edgeOptions)) {
                                _mainForm.browserHandle = GetBrowserHandle();
                                progressForm.Invoke((MethodInvoker)(() => {
                                    progressForm.AbortButton.Enabled = true;
                                }));
                                driver.Navigate().GoToUrl(_mainForm.flashair.Url!);
                                ReadOnlyCollection<IWebElement> elms = driver.FindElements(By.XPath(@"//*[@id='thumbnail']/div"));
                                list = (new List<IWebElement>(elms)).ConvertAll(elm => elm.Text);
                            }
                        } catch (Exception e) {
                            HandleException(progressForm, e);
                            return false;
                        }

                    }
                    //���X�g����ɂ���
                    _mainForm.Invoke((MethodInvoker)(() => {
                        _mainForm.CsvFileListBox.Items.Clear();
                        _mainForm.CsvFileNameLabel.Text = ITEM_NOT_SELECTED;
                    }));

                    if (list.Count == 0) {
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(progressForm,
                                "CSV�t�@�C����1��������܂���ł���\n" +
                                "�w�肵��URL��FlashAir�ł͂Ȃ��\��������܂�");
                        }));
                        return false;
                    }

                    //2023/9/1�����ǉ�
                    //�t�@�C����������(20����?)�Ȃ�ƁA
                    //�u���E�U�ł̓t�@�C�����t���ɃO���[�v�����ꂽ�\���ɂȂ�
                    //�擾����list�̗v�f�ɂ̓t�@�C�����t�����݂���悤�ɂȂ�
                    //(��) "20230901\r\n202309.CSV"
                    //�t�@�C�����ƃt�@�C�����t��ʗv�f�ɕ������Ă���
                    var s = String.Join("\r\n", list);
                    list = s.Split("\r\n", StringSplitOptions.None).ToList();

                    //list �������ōi�荞��
                    //RemoveAll() �ƃ����_���ŏ���
                    list.RemoveAll(s => !Regex.IsMatch(s, @"[0-9]{6}.CSV"));

                    //list ���~���\�[�g
                    //LINQ�ō~���\�[�g
                    list.OrderByDescending(s => s);

                    //���X�g�{�b�N�X�ɒǉ�
                    if (list.Count > 0) {
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.CsvFileListBox.Items.AddRange(list.ToArray());
                            _mainForm.CsvFileListBox.SelectedIndex = 0;
                            _mainForm.CsvFileNameLabel.Text = _mainForm.CsvFileListBox.SelectedItem!.ToString();
                        }));
                    }
                    return true;
                });
            }

            private IntPtr? GetBrowserHandle() {
                IntPtr? browserHandle = null;
                foreach (Process p in Process.GetProcesses()) {
                    if (p.MainWindowTitle.Contains("data:,") == true) {
                        browserHandle = p.MainWindowHandle;
                        break;
                    }
                }
                Debug.WriteLine("GetBrowserHandle()�̖߂�l: " + browserHandle);
                return browserHandle;
            }

            private void HandleException(ProgressForm progressForm, Exception e) {
                switch (e) {
                    case InvalidOperationException:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(
                                progressForm,
                                "�u���E�U�𑀍�ł��܂���ł���\n" +
                                "Selenium.WebDriver�ƃu���E�U�̃o�[�W��������v���Ă��邩�m�F���Ă�������");
                        }));
                        break;
                    case NoSuchWindowException:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(
                                progressForm, "�����𒆒f���܂���");
                        }));
                        break;
                    //case WebDriverArgumentException:
                    case WebDriverException:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(
                                progressForm,
                                "FlashAir�ƒʐM�ł��܂���ł���\n" +
                                "FlashAir��URL�����������m�F���Ă�������");
                        }));
                        break;
                    default:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(progressForm, e);
                        }));
                        break;
                }
            }
        }

        //partial�C���q��t����ProgressForm.cs�ɏ����������ɋL�q
        private partial class ProgressForm : GetFlashairCsv.ProgressForm {
            MainForm _mainForm;
            string _caption;

            public ProgressForm(MainForm mainForm, string caption, ProgressBarStyle progressBarStyle = ProgressBarStyle.Marquee) {
                var systemMenu = GetSystemMenu(this.Handle, false);
                _mainForm = mainForm;
                _caption = caption;

                RemoveMenu(systemMenu, 5, MF_BYPOSITION);
                //����{�^���𖳌���
                RemoveMenu(systemMenu, SC_CLOSE, MF_BYCOMMAND);
                this.AbortButton.Click += AbortButton_Click;
                this.FormClosing += ProgressForm_FormClosing;
                //�\���ʒu�̐ݒ�
                var point = _mainForm.Location;
                this.Bounds = new System.Drawing.Rectangle(
                    point.X + 100, point.Y + 150, this.Size.Width, this.Size.Height);
                this.Text = caption;
                this.ProgressBar.Style = progressBarStyle;
                //�������t�H�[����\��
                //.Show()�Ń��[�h���X�A.ShowDialog()�Ń��[�_��
                this.Show();
                //progressForm �̃��x���������\������
                //�����̈�(��ʍX�V���K�v�ȗ̈�)���ĕ`�悷��
                this.Update();
            }

            private void ProgressForm_FormClosing(object? sender, FormClosingEventArgs e) {
                //�u���E�U��"FlashAir"�̃^�C�g���̂܂܎c��ꍇ�����邽�ߏI��������
                if (_mainForm.browserHandle != null) {
                    SendMessage((IntPtr)_mainForm.browserHandle, WM_CLOSE, 0, 0);
                }
            }

            private void AbortButton_Click(object? sender, EventArgs e) {
                //�u���E�U���I����GoToUrl()�����ŃG���[�𔭐������邱�ƂŒ��f����
                Debug.WriteLine("GoToUrl()�������f��...");
                this.StatusLabel.Text = "���f��...";
                /*
                foreach (Process p in Process.GetProcesses()) {
                    if (p.MainWindowTitle.IndexOf("data:,") >= 0) {
                        //GetWindowRect(p.MainWindowHandle, out RECT rect);
                        //if (rect.left == OUT_OF_SCREEN && rect.top == OUT_OF_SCREEN) {
                        //    p.CloseMainWindow();
                        //    return;
                        //}
                        if (IsIconic(p.MainWindowHandle)) {
                            p.CloseMainWindow();
                            return;
                        }
                    }
                }
                */
                if (_mainForm.browserHandle != null) {
                    Debug.WriteLine("SendMessage()���s");
                    SendMessage((IntPtr)_mainForm.browserHandle, WM_CLOSE, 0, 0);
                }
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public long StartPos { get; set; } = -1;

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public long EndPos { private get; set; }

            public void SetExcelLabelText(long row) {
                this.ExcelLabel.Text = "[Excel] " + row + "�s�ڏ���";
                this.Update();
            }

            public void SetProgressBarValue(long position) {
                int value;
                if (EndPos == StartPos) {
                    value = 100;
                } else {
                    value = (int)((position - StartPos) * 100 / (EndPos - StartPos));
                }
                if (value < this.ProgressBar.Minimum) {
                    return;
                }
                if (value > this.ProgressBar.Maximum) {
                    return;
                }
                this.ProgressBar.Value = value;
                this.CsvLabel.Text = "[CSV] " + EndPos + "�o�C�g���A" + position + "�o�C�g�Ǎ�";
                this.Update();
            }
        }

        private void ShowErrorMessageBox(Exception e) {
            ShowErrorMessageBox(this, e);
        }
        private void ShowErrorMessageBox(Form form, Exception e) {
            CustomMessageBox.Show(form, e.ToString(), CAPTION_ERROR,
               MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowErrorMessageBox(string text) {
            ShowErrorMessageBox(this, text);
        }
        private void ShowErrorMessageBox(Form form, string text) {
            CustomMessageBox.Show(form, text, CAPTION_ERROR,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowOKMessageBox(string text,
            string caption = CAPTION_INFORMATION,
            MessageBoxIcon icon = MessageBoxIcon.Information) {
            CustomMessageBox.Show(this, text, caption,
                MessageBoxButtons.OK, icon);
        }

        private DialogResult ShowOKCancelMessageBox(IWin32Window owner, string text,
            string caption = CAPTION_QUESTION,
            MessageBoxIcon icon = MessageBoxIcon.Warning,
            MessageBoxDefaultButton button = MessageBoxDefaultButton.Button1) {
            var dialogResult = CustomMessageBox.Show(owner, text, caption,
                MessageBoxButtons.OKCancel, icon, button);
            return dialogResult;
        }

        private void WriteInifileButton_Click(object sender, EventArgs e) {
            flashair.WriteUrlToInifile();
            WriteBrowserToInifile();
        }

        private async void UpdateCsvFileListButton_Click(object sender, EventArgs e) {
            UpdateCsvFileListButton.Enabled = false;
            //�������̃t�H�[����\��
            progressForm = new ProgressForm(this, "���X�g�X�V");

            await csvFileList.Update(progressForm);

            //�������̃t�H�[�������
            progressForm.Close();
            UpdateCsvFileListButton.Enabled = true;
        }

        private void CloseButton_Click(object sender, EventArgs e) {
            if (browserHandle != null) {
                Debug.WriteLine("SendMessage()���s");
                SendMessage((IntPtr)browserHandle, WM_CLOSE, 0, 0);
            }
            this.Close();
        }

        private void CsvFileListBox_SelectedIndexChanged(object sender, System.EventArgs e) {
            // Get the currently selected item in the ListBox.
            csvFileList.FileName = CsvFileListBox.SelectedItem!.ToString()!;
        }

        private abstract class ExcelFile {
            public static int GetBottomRownum(string range) {
                var rownumPart = new Regex(@"[0-9]+$");
                return Int32.Parse(rownumPart.Match(range).ToString());

            }

            //Excel�t�@�C���̑��݂��m�F
            //���������ꍇ�͖߂�lNone��Ԃ�
            //������Ȃ��ꍇ�A�V�K�쐬����Ȃ�OK�A�쐬���Ȃ��Ȃ�Cancel��Ԃ�
            protected DialogResult ShowCreateFileDialog(MainForm form, string fileName) {
                //Excel�t�@�C���̑��݂��m�F
                var filePath = Path.GetFullPath(fileName);
                Debug.WriteLine("filePath: " + filePath);
                if (File.Exists(filePath)) {
                    return DialogResult.None;
                }
                //Excel�t�@�C����������Ȃ��ꍇ
                //�V�K�쐬���邩�ǂ����m�F
                return form.ShowOKCancelMessageBox(
                    owner: form,
                    text: EXCEL_FILENAME + " ��������܂���\n�V�K�쐬���܂����H");
            }

            public abstract bool Create(string fileName);
            public abstract bool Open(string fileName);
            public abstract StreamReader? OpenCsvFile(string fileName);
            public abstract Task<int> Write(StreamReader reader);
            public abstract bool ResizeTable();
            public abstract bool ShowInPane();
            public abstract bool ActivateLastCell();
            public abstract bool Close();
        }

        private class ExcelFileUsingOpenXML : ExcelFile {
            //MainForm�̃R���g���[���A�t�B�[���h�̓ǂݏ���
            //���\�b�h�̎��s�̂��߂̕ϐ�
            //���̃R���X�g���N�^�[�̈�������Z�b�g�����
            private readonly MainForm _mainForm;

            private SpreadsheetDocument? _document;
            private WorkbookPart? _workbookPart = null;
            private Sheet? _sheet = null;
            private WorksheetPart? _worksheetPart = null;
            private OOXMLS.Worksheet? _worksheet = null;
            private SheetData? _sheetData = null;
            private SheetView? _sheetView = null;
            private OOXMLS.Pane? _pane = null;
            private string? _tempFileName = null;
            private int _excelRownum = 0;
            private string? _csvDateTime = null;
            private string? _excelDateTime = null;
            private double _excelWh = 0;
            private int _count = 0;

            //�R���X�g���N�^�[
            public ExcelFileUsingOpenXML(MainForm mainForm) {
                _document = null;
                _mainForm = mainForm;
            }

            public int GetCount() {
                return _count;
            }

            public string ConvertDateTimeToString(System.DateTime date, System.DateTime time) {
                return String.Format("{0} {1}",
                    date.ToString(EXCEL_DATE_FORMAT),
                    time.ToString(EXCEL_TIME_FORMAT));
            }

            /// <summary>
            /// ��ԍ�����G�N�Z���̗񖼂𓾂�@(�� 5 �� "E"�j
            /// </summary>
            /// <param name="c"></param>
            /// <returns></returns>
            public string ConvertR1C1ToA1(int r, int c) {
                //A�`Z��܂łȂ炱��1�s��OK
                //return ((char)(64 + (c + 1))).ToString() + r;

                //string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                var s = "";

                for (; c > 0; c = (c - 1) / 26) {
                    var n = (c - 1) % 26;
                    //s = alphabet.Substring(n, 1) + s;
                    s = ((char)(64 + (n + 1))).ToString() + s;
                }
                return s + r;
            }

            public Cell? GetCell(OOXMLS.Worksheet worksheet, string addressName) {
                if ((worksheet == null) || (addressName == null)) {
                    return null;
                }
                //�K�w�\���@Worksheet -> (�q)SheetData -> (��)Row -> (�\��)Cell
                //worksheet.Elements<cell>()�ł͎q�v�f�����擾�ł��Ȃ�
                return worksheet.Descendants<Cell>().
                    FirstOrDefault(c => c.CellReference == addressName);
            }

            private string? GetCellValue(WorkbookPart workbookPart, Cell cell) {
                if ((workbookPart == null) || (cell == null)) {
                    return null;
                }
                var text = cell.InnerText;
                if (cell.DataType == null) {
                    return null;
                }
                switch (cell.DataType!.Value) {
                    case CellValues.SharedString:
                        var index = int.Parse(cell.InnerText);
                        var ssTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (ssTablePart != null) {
                            text = ssTablePart!.SharedStringTable.ElementAt(index).InnerText;
                        }
                        break;
                    case CellValues.Boolean:
                        if (cell.InnerText == "0")
                            text = "FALSE";
                        else
                            text = "TRUE";
                        break;
                }
                return text;
            }

            //[SOLVED]-OPENXML - TABLE CREATION, HOW DO I CREATE TABLES WITHOUT REQUIRING EXCEL TO REPAIR THEM-C#
            //https://www.appsloveworld.com/csharp/100/413/openxml-table-creation-how-do-i-create-tables-without-requiring-excel-to-repai
            public OOXMLS.Table? AppendTable(
                WorksheetPart worksheetPart,
                int rowMin, int rowMax, int colMin, int colMax) {
                if (worksheetPart == null) {
                    return null;
                }
                var tableDefinitionPart =
                    worksheetPart.AddNewPart<TableDefinitionPart>(
                    "rId" + (worksheetPart.TableDefinitionParts.Count() + 1));
                var tableNo = worksheetPart.TableDefinitionParts.Count();

                //�e�[�u���̃Z���͈�
                var reference =
                    ((char)(64 + colMin)).ToString() + rowMin + ":" +
                    ((char)(64 + colMax)).ToString() + rowMax;

                var table = new OOXMLS.Table {
                    Id = (UInt32)tableNo,
                    Name = EXCEL_TABLENAME,
                    DisplayName = EXCEL_TABLENAME,
                    Reference = reference,
                    TotalsRowShown = false
                };

                var autoFilter = new OOXMLS.AutoFilter() { Reference = reference };

                var tableColumns = new TableColumns() {
                    Count = (UInt32)(colMax - colMin + 1)
                };

                //worksheetPart����workbookPart���擾
                var workbookPart = (WorkbookPart?)worksheetPart.GetParentParts().FirstOrDefault(p => p is WorkbookPart);
                for (int i = 0; i < (colMax - colMin + 1); i++) {
                    tableColumns.Append(new TableColumn() {
                        Id = (UInt32)(colMin + i),
                        //Name�́A�ݒ�Ώۂ̃Z���Ɋi�[����Ă���l�Ɠ������e��ݒ肷��
                        Name = GetCellValue(workbookPart!,
                            GetCell(worksheetPart.Worksheet, ConvertR1C1ToA1(rowMin, i + 1))!)
                    });
                }

                TableStyleInfo tableStyleInfo = new TableStyleInfo {
                    Name = EXCEL_TABLESTYLENAME,
                    ShowFirstColumn = false,
                    ShowLastColumn = false,
                    ShowRowStripes = true,
                    ShowColumnStripes = false
                };

                table.Append(autoFilter);
                table.Append(tableColumns);
                table.Append(tableStyleInfo);

                tableDefinitionPart.Table = table;

                var tableParts =
                    //(TableParts?)worksheetPart.Worksheet.ChildElements.Where(
                    //ce => ce is TableParts).FirstOrDefault();
                    // Add table parts only once
                    (TableParts?)worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
                if (tableParts is null) {
                    tableParts = new TableParts();
                    tableParts.Count = (UInt32)0;
                    worksheetPart.Worksheet.Append(tableParts);
                }
                // is not null �� != null �ɂ���ƃG���[
                if (tableParts.Count is not null) {
                    tableParts.Count += (UInt32)1;
                }
                var tablePart = new TablePart { Id = "rId" + tableNo };

                tableParts.Append(tablePart);

                return table;
            }

            public DialogResult ShowCreateFileDialog(string fileName) {
                return base.ShowCreateFileDialog(_mainForm, fileName);
            }

            ////Dynamic comments in Excel using Open XML
            //https://social.msdn.microsoft.com/Forums/en-US/c4400c1f-e4b4-43ed-b037-2f531274ea78/dynamic-comments-in-excel-using-open-xml?forum=exceldev
            public bool InsertComments(WorksheetPart worksheetPart,
                List<string> ColumnName, List<string> CellIndex, List<string> NewCommentList) {
                if ((ColumnName.Count == 0) || (CellIndex.Count == 0) || (NewCommentList.Count == 0)) {
                    return false;
                }
                try {
                    var commentsVmlXml = string.Empty;
                    // Create all the comment VML Shape XML
                    for (var i = 0; i < ColumnName.Count; i++) {
                        commentsVmlXml += GetCommentVMLShapeXML(ColumnName[i], CellIndex[i]);
                    }
                    var vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
                    using (var writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8)) {
                        writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n" +
                        "</o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\r\n  path=\"m,l,21600r21600,l21600,xe\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n </v:shapetype>"
                        + commentsVmlXml + "</xml>");
                    }
                    // Create the comment elements
                    for (var j = 0; j < NewCommentList.Count; j++) {
                        var worksheetCommentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();
                        // We only want one legacy drawing element per worksheet for comments
                        if (worksheetPart.Worksheet.Descendants<LegacyDrawing>().SingleOrDefault() == null) {
                            string vmlPartId = worksheetPart.GetIdOfPart(vmlDrawingPart);
                            var legacyDrawing = new LegacyDrawing() { Id = vmlPartId };
                            worksheetPart.Worksheet.Append(legacyDrawing);
                        }
                        OOXMLS.Comments comments;
                        bool appendComments = false;
                        if (worksheetPart.WorksheetCommentsPart!.Comments != null) {
                            comments = worksheetPart.WorksheetCommentsPart.Comments;
                        } else {
                            comments = new OOXMLS.Comments();
                            appendComments = true;
                        }
                        // We only want one Author element per Comments element
                        if (worksheetPart.WorksheetCommentsPart.Comments == null) {
                            var authors = new Authors();
                            var author = new OOXMLS.Author();
                            author.Text = "Author";
                            authors.Append(author);
                            comments.Append(authors);
                        }
                        OOXMLS.CommentList commentList;
                        bool appendCommentList = false;
                        if ((worksheetPart.WorksheetCommentsPart.Comments != null) &&
                            (worksheetPart.WorksheetCommentsPart.Comments.
                            Descendants<OOXMLS.CommentList>().SingleOrDefault() != null)) {
                            commentList =
                                worksheetPart.WorksheetCommentsPart.Comments!.
                                Descendants<OOXMLS.CommentList>().Single();
                        } else {
                            commentList = new OOXMLS.CommentList();
                            appendCommentList = true;
                        }
                        var comment = new OOXMLS.Comment() {
                            Reference = ColumnName[j] + CellIndex[j],
                            AuthorId = (UInt32Value)0U
                        };
                        var commentTextElement = new CommentText();
                        var run = new OOXMLS.Run();
                        var runProperties = new OOXMLS.RunProperties();
                        runProperties.Append(new OOXMLS.Bold());
                        runProperties.Append(new OOXMLS.FontSize() { Val = 9D });
                        runProperties.Append(new OOXMLS.Color() { Indexed = (UInt32Value)81U });
                        //runProperties.Append(new RunFont() { Val = "MS P �S�V�b�N" });
                        runProperties.Append(new RunFont() { Val = "�l�r �o�S�V�b�N" });
                        runProperties.Append(new OOXMLS.FontFamily() { Val = 3 });
                        runProperties.Append(new RunPropertyCharSet() { Val = 128 });
                        run.Append(runProperties);
                        run.Append(new OOXMLS.Text() { Text = NewCommentList[j] });
                        commentTextElement.Append(run);
                        comment.Append(commentTextElement);
                        commentList.Append(comment);
                        // Only append the Comment List if this is the first time adding a comment
                        if (appendCommentList) {
                            comments.Append(commentList);
                        }
                        // Only append the Comments if this is the first time adding Comments
                        if (appendComments) {
                            worksheetCommentsPart.Comments = comments;
                        }
                    }
                } catch (Exception e) {
                    _mainForm.ShowErrorMessageBox(e);
                    return false;
                }
                return true;
            }

            /// <summary>
            /// Creates the VML Shape XML for a comment. It determines the positioning of the
            /// comment in the excel document based on the column name and row index.
            /// </summary>
            /// <param name="columnName">Column name containing the comment</param>
            /// <param name="rowIndex">Row index containing the comment</param>
            /// <returns>VML Shape XML for a comment</returns>
            private string GetCommentVMLShapeXML(string columnName, string rowIndex) {
                var commentVmlXml = string.Empty;

                // Parse the row index into an int so we can subtract one
                int commentRowIndex;
                if (int.TryParse(rowIndex, out commentRowIndex) == false) {
                    return commentVmlXml;

                }
                commentRowIndex -= 1;
                commentVmlXml = "<v:shape id=\"" + Guid.NewGuid().ToString().Replace("-", "") + "\" type=\"#_x0000_t202\" style=\'position:absolute;\r\n  margin-left:59.25pt;margin-top:1.5pt;width:96pt;height:55.5pt;z-index:1;\r\n  visibility:hidden\' fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\r\n  <v:fill color2=\"#ffffe1\"/>\r\n" +
                    "<v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\r\n  <v:path o:connecttype=\"none\"/>\r\n  <v:textbox style=\'mso-fit-shape-to-text:true'>\r\n   <div style=\'text-align:left\'></div>\r\n  </v:textbox>\r\n  <x:ClientData ObjectType=\"Note\">\r\n   <x:MoveWithCells/>\r\n" +
                    "<x:SizeWithCells/>\r\n   <x:Anchor>\r\n" + GetAnchorCoordinatesForVMLCommentShape(columnName, rowIndex) + "</x:Anchor>\r\n   <x:AutoFill>False</x:AutoFill>\r\n   <x:Row>" + commentRowIndex + "</x:Row>\r\n   <x:Column>" + GetColumnIndexFromName(columnName) + "</x:Column>\r\n  </x:ClientData>\r\n </v:shape>";
                return commentVmlXml;
            }

            /// <summary>
            /// Gets the coordinates for where on the excel spreadsheet to display the VML comment shape
            /// </summary>
            /// <param name="columnName">Column name of where the comment is located (ie. B)</param>
            /// <param name="rowIndex">Row index of where the comment is located (ie. 2)</param>
            /// <returns><see cref="<x:Anchor>"/> coordinates in the form of a comma separated list</returns>
            private string GetAnchorCoordinatesForVMLCommentShape(string columnName, string rowIndex) {
                var coordinates = string.Empty;
                //int startingRow = 0;
                var startingColumn = GetColumnIndexFromName(columnName)!.Value;

                // From (upper right coordinate of a rectangle)
                // [0] Left column
                // [1] Left column offset
                // [2] Left row
                // [3] Left row offset
                // To (bottom right coordinate of a rectangle)
                // [4] Right column
                // [5] Right column offset
                // [6] Right row
                // [7] Right row offset
                var coordList = new List<int>(8) { 0, 0, 0, 0, 0, 0, 0, 0 };

                if (int.TryParse(rowIndex, out int startingRow) == false) {
                    return coordinates;
                }

                // Make the row be a zero based index
                startingRow -= 1;

                coordList[0] = startingColumn + 1; // If starting column is A, display shape in column B
                coordList[1] = 15;
                coordList[2] = startingRow;
                coordList[4] = startingColumn + 3; // If starting column is A, display shape till column D
                coordList[5] = 15;
                coordList[6] = startingRow + 3; // If starting row is 0, display 3 rows down to row 3

                // The row offsets change if the shape is defined in the first row
                if (startingRow == 0) {
                    coordList[3] = 2;
                    coordList[7] = 16;
                } else {
                    coordList[3] = 10;
                    coordList[7] = 4;
                }
                coordinates = string.Join(",", coordList.ConvertAll<string>(x => x.ToString()).ToArray());
                return coordinates;
            }

            /// <summary>
            /// Given just the column name (no row index), it will return the zero based column index.
            /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
            /// A length of three can be implemented when needed.
            /// </summary>
            /// <param name="columnName">Column Name (ie. A or AB)</param>
            /// <returns>Zero based index if the conversion was successful; otherwise null</returns>
            public static int? GetColumnIndexFromName(string columnName) {
                int? columnIndex = null;

                string[] colLetters = Regex.Split(columnName, "([A-Z]+)");
                colLetters = colLetters.Where(s => !string.IsNullOrEmpty(s)).ToArray();
                var Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J' };

                if (colLetters.Length <= 2) {
                    int index = 0;
                    foreach (string col in colLetters) {
                        var col1 = colLetters.ElementAt(index).ToCharArray().ToList();
                        var indexValue = Letters.IndexOf(col1.ElementAt(index));

                        if (indexValue != -1) {
                            // The first letter of a two digit column needs some extra calculations
                            if ((index == 0) && (colLetters.Count() == 2)) {
                                columnIndex = (columnIndex == null) ?
                                    (indexValue + 1) * 26 :
                                    columnIndex + ((indexValue + 1) * 26);
                            } else {
                                columnIndex = (columnIndex == null) ?
                                    indexValue :
                                    columnIndex + indexValue;
                            }
                        }
                        index++;
                    }
                }
                return columnIndex;
            }

            public override bool Create(string fileName) {
                if (String.IsNullOrEmpty(fileName)) {
                    return false;
                }

                try {
                    _document = SpreadsheetDocument.
                        Create(fileName, SpreadsheetDocumentType.Workbook);
                    var workbookPart = _document.AddWorkbookPart();
                    workbookPart.Workbook = new OOXMLS.Workbook();
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var worksheets = _document.WorkbookPart!.Workbook.AppendChild<OOXMLS.Sheets>(new OOXMLS.Sheets());
                    var sheet = new Sheet() {
                        Id = _document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = EXCEL_SHEETNAME
                    };
                    worksheets.Append(sheet);

                    //�V�[�g�̍쐬
                    var sheetData = new SheetData();

                    //A��̍쐬
                    var row = new OOXMLS.Row() {
                        RowIndex = 1U,
                        Spans = new ListValue<OOXML.StringValue>(),
                        Height = 36U,
                        CustomHeight = true
                    };

                    //A1�Z���̍쐬
                    Cell cell;
                    cell = new Cell() {
                        CellReference = "A1",
                        //R1C1�`���ł��B�������Q�Ƃ̍ۂ͍쐬�����`���ł����Q�Ƃł��Ȃ�
                        //Excel�ň�x�ۑ���������A1�`���ɕϊ������l�q
                        DataType = CellValues.String,
                        CellValue = new OOXMLS.CellValue(EXCEL_HEADER_IDENTIFIER),
                    };
                    row.Append(cell);
                    sheetData.Append(row);

                    var worksheet = new OOXMLS.Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                    //worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    //worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                    worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                    //�f�[�^�͈͂̐ݒ�
                    var sheetDimension =
                        new SheetDimension() { Reference = "A1:A1" };

                    //�X�^�C���V�[�g�̒ǉ�
                    var stylesPart =
                        workbookPart.AddNewPart<WorkbookStylesPart>();
                    var stylesheet = new OOXMLS.Stylesheet();

                    //�t�H���g��`
                    var fonts = new OOXMLS.Fonts() { Count = 1 };
                    var font = new OOXMLS.Font() {
                        FontSize = new OOXMLS.FontSize() { Val = 11 },
                        FontName = new FontName() { Val = new StringValue("���S�V�b�N") },
                        FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                        FontCharSet = new OOXMLS.FontCharSet() { Val = 128 },
                        FontScheme = new OOXMLS.FontScheme() { Val = FontSchemeValues.Minor }
                    };
                    fonts.Append(font);
                    stylesheet.Append(fonts);

                    //�h��Ԃ��̒�`
                    var fills = new OOXMLS.Fills();
                    fills.Append(new OOXMLS.Fill() {
                        PatternFill = new OOXMLS.PatternFill {
                            PatternType = PatternValues.None
                        }
                    });
                    fills.Append(new OOXMLS.Fill() {
                        PatternFill = new OOXMLS.PatternFill {
                            PatternType = PatternValues.Gray125
                        }
                    });
                    stylesheet.Append(fills);

                    //�{�[�_�[��`
                    var borders = new OOXMLS.Borders() { Count = 1 };
                    var border = new OOXMLS.Border();
                    border.Append(new OOXMLS.LeftBorder());
                    border.Append(new OOXMLS.RightBorder());
                    border.Append(new OOXMLS.TopBorder());
                    border.Append(new OOXMLS.BottomBorder());
                    border.Append(new OOXMLS.DiagonalBorder());
                    borders.Append(border);
                    borders.Append(new OOXMLS.Border());

                    stylesheet.Append(borders);

                    //�Z���X�^�C���t�H�[�}�b�g��`
                    var cellStyleFormats = new OOXMLS.CellStyleFormats() { Count = 1U };
                    cellStyleFormats.Append(new OOXMLS.CellFormat() {
                        NumberFormatId = 0,
                        FontId = 0,
                        FillId = 0,
                        BorderId = 0,
                        Alignment = new OOXMLS.Alignment {
                            Vertical = OOXMLS.VerticalAlignmentValues.Center
                        }
                    });
                    stylesheet.Append(cellStyleFormats);

                    //�Z���t�H�[�}�b�g��`��ǉ�
                    var cellFormats = new OOXMLS.CellFormats();
                    //�W���p
                    cellFormats.Append(new OOXMLS.CellFormat() {
                        FormatId = 0,
                        NumberFormatId = 0,
                        FontId = (UInt32Value)0U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        Alignment = new OOXMLS.Alignment() {
                            Vertical = OOXMLS.VerticalAlignmentValues.Center
                        }
                    });
                    //�����܂�Ԃ��p
                    cellFormats.Append(new OOXMLS.CellFormat() {
                        FormatId = 0,
                        NumberFormatId = 0,
                        FontId = (UInt32Value)0U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        ApplyAlignment = true,
                        Alignment = new OOXMLS.Alignment() {
                            Vertical = OOXMLS.VerticalAlignmentValues.Center,
                            WrapText = true
                        }
                    });
                    cellFormats.Count
                      = new UInt32Value((uint)cellFormats.Count());
                    stylesheet.Append(cellFormats);

                    stylesPart.Stylesheet = stylesheet;

                    //�V�[�g�r���[�̍쐬
                    var sheetViews = new OOXMLS.SheetViews();
                    SheetView sheetView = new SheetView() {
                        TabSelected = true,
                        WorkbookViewId = (UInt32Value)0U
                    };

                    //�E�B���h�E�g�̌Œ�̐ݒ�
                    //selection1�����Append����ƃG���[�ɂȂ�
                    var pane = new OOXMLS.Pane() {
                        //HorizontalSplit = 1, //��Œ�
                        VerticalSplit = 1,     //�s�Œ�
                        TopLeftCell = "A2",    //���̃y�C���̍���̃Z��
                        ActivePane = PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen
                    };
                    sheetView.Append(pane);

                    //�A�N�e�B�u�Z���̐ݒ�
                    var selection = new OOXMLS.Selection() {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = "A1",
                        SequenceOfReferences = new ListValue<OOXML.StringValue>() {
                            InnerText = "A1" //R1C1�`���͕s��
                        }
                    };
                    sheetView.Append(selection);
                    sheetViews.Append(sheetView);

                    worksheet.Append(sheetDimension);
                    worksheet.Append(sheetViews);

                    //�y�[�W�ݒ�
                    var sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
                    var pageMargins = new OOXMLS.PageMargins() {
                        Left = 0.7D,
                        Right = 0.7D,
                        Top = 0.75D,
                        Bottom = 0.75D,
                        Header = 0.3D,
                        Footer = 0.3D
                    };
                    worksheet.Append(sheetFormatProperties);

                    //�񕝂̐ݒ�
                    //sheetData1�����Append����ƃG���[�ɂȂ�
                    var lstColumns = new OOXMLS.Columns();
                    lstColumns.Append(new OOXMLS.Column() {
                        Min = 1,
                        Max = 1,
                        Width = 15,
                        CustomWidth = true
                    });
                    lstColumns.Append(new OOXMLS.Column() {
                        Min = 2,
                        Max = 11,
                        Width = 10,
                        CustomWidth = true
                    });
                    worksheet.Append(lstColumns);

                    worksheet.Append(sheetData);
                    worksheetPart.Worksheet = worksheet;

                    //�R�����g�̑}��
                    //Excel�ŊJ���ăR�����g�̕ҏW��ԂŖ��O�{�b�N�X�ɕ\�������
                    //���O(����16�i�����ݒ肳��Ă���)��ݒ肷����@��������Ȃ�����
                    //��������Excel�ŃR�����g�̑}�������Ă����O�̕ύX�͂ł��Ȃ�
                    List<string> ColumnName = new List<string>() { "A" };
                    List<string> CellIndex = new List<string>() { "1" };
                    List<string> NewCommentList = new List<string>() { "�V�[�g���ʂ̂��ߕύX�s��" };
                    InsertComments(worksheetPart, ColumnName, CellIndex, NewCommentList);

                    _document.Save();
                    _document.Close();

                    Debug.WriteLine("ExcelFileUsingOpenXML.Create() ����");
                    return true;
                } catch (Exception e) {
                    Debug.WriteLine(e);
                    _mainForm.ShowErrorMessageBox(fileName + " ��V�K�쐬�ł��܂���ł���");
                    return false;
                }
            }

            //Open()�ɂ���
            //Open()������Dispose()���Ȃ��ƃt�@�C�����j������
            //�������t�@�C��(�̃^�C���X�^���v�H)���X�V����Ă��܂�
            //��ؕۑ������I�u�W�F�N�g��j��������@���Ȃ��H
            //�������@
            //Open()�ɂď����@�t�@�C�����J����
            //�ꎞ�t�@�C�������ꎞ�t�@�C����ōX�V����
            //Close()�ɂď����@�ۑ���
            //���̃t�@�C���ɏ㏑�����Ĉꎞ�t�@�C�����폜����
            public override bool Open(string fileName) {
                if (String.IsNullOrEmpty(fileName)) {
                    return false;
                }

                try {
                    try {
                        _tempFileName = Path.GetTempFileName();
                        Debug.WriteLine("tempFileName: " + _tempFileName);
                        //tempFileName: C:\Users\(���[�U�[��)\AppData\Local\Temp\tmp????.tmp
                    } catch (Exception) {
                        _mainForm.ShowErrorMessageBox("�ꎞ�t�@�C�����쐬�ł��܂���ł���");
                        return false;
                    }
                    try {
                        File.Copy(EXCEL_FILENAME, _tempFileName, true);
                    } catch {
                        _mainForm.ShowErrorMessageBox(
                            "Excel�t�@�C�����ꎞ�t�@�C���փR�s�[�ł��܂���ł���");
                        return false;
                    }
                    _document = SpreadsheetDocument.Open(_tempFileName, true,
                       new OpenSettings { AutoSave = false });
                } catch (System.IO.IOException) {
                    _mainForm.ShowErrorMessageBox(EXCEL_FILENAME + "�ɃA�N�Z�X�ł��܂���");
                    return false;
                } catch (Exception e) {
                    _mainForm.ShowErrorMessageBox(e);
                    return false;
                }

                //Excel�t�@�C���̒��g�����������m�F
                _workbookPart = _document.WorkbookPart;
                var sheets =
                    _workbookPart!.Workbook.GetFirstChild<OOXMLS.Sheets>();
                _sheet =
                    sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == EXCEL_SHEETNAME);
                if (_sheet == null) {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox(EXCEL_SHEETNAME + "��������܂���");
                    return false;
                }
                Debug.WriteLine("sheet.Name: " + _sheet.Name);
                _worksheetPart = (WorksheetPart)_workbookPart.GetPartById(_sheet.Id!);
                _worksheet = _worksheetPart.Worksheet;
                var a1 = GetCell(_worksheet, "A1");
                if (a1 == null) {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox("A1�Z����������܂���");
                    return false;
                }
                var cellValue = GetCellValue(_workbookPart, a1);
                Debug.WriteLine("A1: " + cellValue);
                if (cellValue != EXCEL_HEADER_IDENTIFIER) {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox(
                        "30���l�f�[�^��������܂���\n" +
                        "A1�Z����" + EXCEL_HEADER_IDENTIFIER);
                    return false;
                }

                //�f�[�^�̃Z���͈͂��擾
                var usedRange = _worksheetPart.Worksheet.SheetDimension!.Reference;
                Debug.WriteLine("usedRange: " + usedRange);

                //�ŏI�s���擾
                _excelRownum = GetBottomRownum(usedRange!);
                Debug.WriteLine("excelRownum: " + _excelRownum);

                //�ŏI�s�̃f�[�^���擾
                if (_excelRownum == 1) {
                    //�w�b�_�[�s�����Ȃ��ꍇ
                    _excelDateTime = "";
                    _excelWh = 0.0;
                } else {
                    var date = GetCell(_worksheet, "A" + _excelRownum);
                    var time = GetCell(_worksheet, "B" + _excelRownum);
                    var wh = GetCell(_worksheet, "C" + _excelRownum);
                    try {
                        _excelDateTime = "";
                        _excelWh = 0.0;
                        //�蓮�Ńw�b�_�[�s�ȊO��S�폜���Ă���ꍇ�̑Ώ�
                        //�e�[�u�����ݒ肳��Ă���ƌ����ڂ̓w�b�_�[�s�����Ȃ��Ă�
                        //2�s�ڂ�Row�I�u�W�F�N�g��Cell�I�u�W�F�N�g(InnerText��"")�����݂�
                        //_excelRownum�̒l��2�ƔF�������
                        if (date!.InnerText == "") {
                            _excelRownum = 1;
                        } else {
                            var dateSerial = double.Parse(date.InnerText);
                            var timeSerial = double.Parse(time!.InnerText);
                            _excelDateTime = ConvertDateTimeToString(
                                System.DateTime.FromOADate(dateSerial),
                                System.DateTime.FromOADate(timeSerial));
                            // �����̍Ō�̃f�[�^�܂ŏ���������Ă��Ȃ��̂Ɏ��̌��������݂��悤�Ƃ����ꍇ�Ɋm�F
                            var dateTime = System.DateTime.FromOADate(dateSerial + timeSerial);
                            dateTime = dateTime.AddMinutes(30);
                            if ((dateTime.Day != 1) || (dateTime.Hour != 1) || (dateTime.Minute != 1)) {
                                var csvFileName = _mainForm.CsvFileNameLabel.Text;
                                if ((dateTime.Year != int.Parse(csvFileName[..4])) ||
                                    (dateTime.Month != int.Parse(csvFileName[4..6]))) {
                                    DialogResult dialogResult = DialogResult.None;
                                    _mainForm.Invoke((MethodInvoker)(() => {
                                        dialogResult = _mainForm.ShowOKCancelMessageBox(
                                           owner: _mainForm.progressForm!,
                                           text: string.Format(
                                               "{0}�N{1}�������Ō�܂ŏ������݂���Ă��܂���\n\n" +
                                               "�������݂𑱍s���܂����H", dateTime.Year, dateTime.Month),
                                           button: MessageBoxDefaultButton.Button2);
                                    }));
                                    if (dialogResult == DialogResult.Cancel) {
                                        _document.Dispose();
                                        File.Delete(_tempFileName!);
                                        return false;
                                    }
                                }
                            }
                            _excelWh = Convert.ToDouble(wh!.InnerText);
                            Debug.WriteLine("ExcelFileUsingOpenXML._excelDateTime: " + _excelDateTime);
                        }
                    } catch (Exception e) {
                        _document.Dispose();
                        File.Delete(_tempFileName);
                        _mainForm.ShowErrorMessageBox(e);
                        return false;
                    }
                    _mainForm.ExcelLastDataLabel.Text = string.Format("{0}�s {1} {2}kWh",
                        _excelRownum, _excelDateTime, (int)((_excelWh + 500) / 1000));
                }
                Debug.WriteLine("ExcelFileUsingOpenXML.Open() ����");
                return true;
            }

            public override StreamReader? OpenCsvFile(string fileName) {
                if ((_document == null) || (String.IsNullOrEmpty(fileName))) {
                    return null;
                }
                if (fileName == ITEM_NOT_SELECTED) {
                    return null;
                }

                StreamReader reader;
                try {
                    var csvFilePath =
                        Path.Combine(Application.StartupPath, fileName);
                    Debug.WriteLine("csvFilePath: " + csvFilePath);
                    reader = new StreamReader(csvFilePath, Encoding.GetEncoding("UTF-8"));
                } catch (FileNotFoundException) {
                    _document.Dispose();
                    File.Delete(_tempFileName!);
                    _mainForm.ShowErrorMessageBox("CSV�t�@�C����������܂���");
                    return null;
                } catch (Exception e) {
                    _document.Dispose();
                    _mainForm.ShowErrorMessageBox(e);
                    return null;
                }

                //CSV�t�@�C���̃w�b�_�[��ǂݍ���
                string line = reader.ReadLine()!;
                if (line == null) {
                    _document.Dispose();
                    _mainForm.ShowErrorMessageBox("CSV�t�@�C���̒��g����ł�");
                    return null;
                }
                string[] cols = line.Split(',');
                //Excel�t�@�C���̃w�b�_�[�����w�b�_�[��������
                var b1 = GetCell(_worksheet!, "B1");
                if (b1 == null) {
                    //CSV�t�@�C���̃w�b�_�[���R�s�[
                    //�w�b�_�[�̃Z������
                    //�Z���������ǉ��ς��m�F
                    var stylesPart = _workbookPart!.WorkbookStylesPart;
                    var cellFormats = stylesPart!.Stylesheet.CellFormats;
                    object? headerStyleIndex = null;
                    UInt32Value index = 0;
                    foreach (OOXMLS.CellFormat f in cellFormats!) {
                        var alignment = f.Elements<OOXMLS.Alignment>().FirstOrDefault();
                        if ((alignment != null) &&
                            (alignment.WrapText is not null) &&
                            (alignment.WrapText == true)) {
                            headerStyleIndex = new UInt32Value(index);
                            break;

                        }
                        index++;
                    }
                    Debug.WriteLine("headerStyleIndex: " + headerStyleIndex);

                    var row = _worksheet!.Descendants<OOXMLS.Row>().
                            Where(r => r.RowIndex! == 1).First();
                    Cell? cell;
                    string cellReference = "";
                    for (int n = 0; n < cols.Length; n++) {
                        cellReference = ConvertR1C1ToA1(1, n + 1);
                        cell = GetCell(_worksheet, cellReference);
                        if (cell == null) {
                            cell = new() {
                                CellReference = cellReference,
                                DataType = CellValues.String,
                                CellValue = new OOXMLS.CellValue(cols[n]),
                            };
                            row.Append(cell);
                        }
                        if (headerStyleIndex != null) {
                            cell.StyleIndex = (UInt32Value)headerStyleIndex;
                        }
                    }
                }
                Debug.WriteLine("ExcelFileUsingOpenXML.OpenCsvFile() ����");
                return reader;
            }

            public override async Task<int> Write(StreamReader reader) {
                return await Task.Run(() => {
                    if ((_document == null) || (reader == null)) {
                        return ERROR_RETURN_VALUE;
                    }
                    //Excel�X�^�C���V�[�g�ɃZ���t�H�[�}�b�g��`��ǉ�
                    //Cell�̏�����ݒ肷��(OpenXML��)
                    //https://www.cloverfield.co.jp/2020/02/28/cell%E3%81%AE%E6%9B%B8%E5%BC%8F%E3%82%92%E8%A8%AD%E5%AE%9A%E3%81%99%E3%82%8Bopenxml%E7%B7%A8/
                    var stylesPart = _workbookPart!.WorkbookStylesPart;
                    var cellFormats = stylesPart!.Stylesheet.CellFormats;

                    //���t�̃Z������
                    ////NumberFormatId=14�̑g���ݏ��� [���t]��� *2012/3/14
                    //�Z���������ǉ��ς��m�F
                    object? dateStyleIndex = null;
                    UInt32Value index = 0;
                    foreach (OOXMLS.CellFormat f in cellFormats!) {
                        if (f.NumberFormatId! == 14) {
                            dateStyleIndex = new UInt32Value(index);
                            break;
                        }
                        index++;
                    }
                    Debug.WriteLine("dateStyleIndex: " + dateStyleIndex);
                    //�Z��������ǉ�
                    if (dateStyleIndex == null) {
                        stylesPart.Stylesheet.CellFormats!.AppendChild(new OOXMLS.CellFormat() {
                            FormatId = 0,
                            NumberFormatId = 14, //mm-dd-yy
                            FontId = (UInt32Value)0U,
                            FillId = (UInt32Value)0U,
                            BorderId = (UInt32Value)0U,
                            ApplyNumberFormat = OOXML.BooleanValue.FromBoolean(true),
                            ApplyAlignment = true,
                            Alignment = new OOXMLS.Alignment {
                                Vertical = OOXMLS.VerticalAlignmentValues.Center
                            }
                        });
                        stylesPart.Stylesheet.CellFormats.Count
                          = new UInt32Value((uint)stylesPart.Stylesheet.CellFormats.Count());
                        dateStyleIndex =
                            new UInt32Value(stylesPart.Stylesheet.CellFormats.Count - 1);
                    }

                    //�����̃Z������
                    //Apply OpenXML excel number formatcode to string value in C#
                    //https://stackoverflow.com/questions/25228471/apply-openxml-excel-number-formatcode-to-string-value-in-c-sharp
                    //https://learn.microsoft.com/ja-jp/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
                    //NumberFormatId=20�̑g���ݏ���"h:mm"�́A[����]���13:30�ɂł��Ȃ�
                    //FormatCode="h:mm;@"�̃��[�U�[��`��ǉ�����K�v������
                    //���[�U�[��`��NumberFormatId��176�ȍ~�炵��
                    //176�������l�ɂ��ď��ԂŌ��Ă����ċ󂢂Ă����甭�Ԃ��邱�Ƃɂ���
                    //�G���[�͏o�Ă��Ȃ����A���̕��@�őS�����Ȃ��̂��s��
                    //<NumberFormatId�̔��Ԃ̎d�g�݂�����>
                    //�Z���ɑ΂��đg���݂�[����]���13:30��Excel�ő��삵�Đݒ肷���
                    //[����]���13:30�ɂȂ邪�Astyles.xml��`���ƃ��[�U�[��`�ɂȂ�
                    //FormatCode��"h:mm;@"��,NumberFormatId�̔��Ԃ�
                    //176�͋󂢂Ă���̂�176�ɂȂ炸�A���x��180,181�ƈ�肵�Ȃ�
                    // �t�@�C���ł͂Ȃ�Excel���O���[�o���ɊǗ����Ă���?(PC���ʂȂ�Ǘ��͖���)
                    // �����_���ɔԍ����쐬���ăt�@�C�����Ŗ��g�p�Ȃ猈��?
                    // 2�̃t�@�C��(����Id��Code�̓��e���Ⴄ)��1�ɂ������ǂ��Ȃ�̂�?

                    // "h:mm;@"�̃��[�U�[��`���o�^�ς��m�F
                    stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
                    var timeFormatCode =
                        OOXML.StringValue.FromString(EXCEL_STYLE_TIME_FORMATCODE);
                    var numberingFormats = stylesPart.Stylesheet.NumberingFormats;
                    OOXMLS.NumberingFormat numberingFormat;
                    var tempTimeFormat =
                        numberingFormats.Elements<OOXMLS.NumberingFormat>().FirstOrDefault(f => f.FormatCode == timeFormatCode);
                    if (tempTimeFormat != null) {
                        numberingFormat = tempTimeFormat;
                    } else {
                        numberingFormat = new();
                        var tempUserFormat = numberingFormats.
                            Elements<OOXMLS.NumberingFormat>().
                            Where(f => f.NumberFormatId! >= 176).LastOrDefault();
                        if (tempUserFormat != null) {
                            numberingFormat.NumberFormatId =
                                tempUserFormat.NumberFormatId! + 1;
                        } else {
                            numberingFormat.NumberFormatId =
                            UInt32Value.FromUInt32(176);
                            //UInt32Value.FromUInt32(iExcelIndex++);
                        }
                        numberingFormat.FormatCode = timeFormatCode;
                        stylesPart.Stylesheet.NumberingFormats.Append(numberingFormat);
                    }
                    Debug.WriteLine("numberingFormat.NumberFormatId: " + numberingFormat.NumberFormatId);

                    //�Z���������ǉ��ς��m�F
                    object? timeStyleIndex = null;
                    index = 0;
                    foreach (OOXMLS.CellFormat f in cellFormats) {
                        if (f.NumberFormatId! == numberingFormat.NumberFormatId!) {
                            timeStyleIndex = new UInt32Value(index);
                            break;
                        }
                        index++;
                    }
                    Debug.WriteLine("timeStyleIndex: " + timeStyleIndex);

                    //�Z��������ǉ�
                    if (timeStyleIndex! == null) {
                        //�Z���������쐬
                        stylesPart.Stylesheet.CellFormats!.AppendChild(new OOXMLS.CellFormat() {
                            FormatId = (UInt32Value)0U,
                            NumberFormatId = numberingFormat.NumberFormatId,
                            FontId = (UInt32Value)0U,
                            FillId = (UInt32Value)0U,
                            BorderId = (UInt32Value)0U,
                            ApplyNumberFormat = OOXML.BooleanValue.FromBoolean(true),
                            ApplyAlignment = true,
                            Alignment = new OOXMLS.Alignment {
                                Vertical = OOXMLS.VerticalAlignmentValues.Center
                            }
                        });
                        stylesPart.Stylesheet.CellFormats.Count
                          = new UInt32Value((uint)stylesPart.Stylesheet.CellFormats.Count());
                        timeStyleIndex = new UInt32Value(stylesPart.Stylesheet.CellFormats.Count - 1);
                    }

                    //Excel�t�@�C���ɏ�������
                    OOXMLS.Row? row = null;
                    Cell? date = null;
                    Cell? time = null;
                    Cell? wh = null;
                    Cell? cell = null;
                    string[] cols = { "" };
                    string line = "";
                    string prevCsvDateTime = _excelDateTime!;
                    bool dontShowAgain = false;
                    while (reader.Peek() >= 0) {
                        //���Ԃ������鏈���ł́u�����Ȃ��v���������ɂ́H
                        //https://atmarkit.itmedia.co.jp/ait/articles/0403/19/news088.html
                        //Application.DoEvents();

                        // �ǂݍ��񂾕����񂩂琧�䕶��������
                        // ��̍s�Ȃ珈�����X�L�b�v
                        line = string.Concat(reader.ReadLine()!.Where(c => !char.IsControl(c)));
                        if (line == "") {
                            continue;
                        }
                        // ��������J���}��؂�Ŕz��Ɋi�[
                        cols = line.Split(',');
                        _csvDateTime = cols[0] + " " + cols[1];
                        //_csvDateTime�̏���: yyyy/mm/dd hh:mm
                        //Excel�ŏI�s�̓������������������O�Ȃ�X�L�b�v
                        if ((System.DateTime.Parse(_csvDateTime) - System.DateTime.Parse(_excelDateTime!)).TotalMinutes <= 0) {
                            continue;
                        }
                        //�O��̃��[�v�ŏ��������f�[�^�̓����Ƃ̍���30�����傫����Ίm�F
                        if ((System.DateTime.Parse(_csvDateTime) - System.DateTime.Parse(prevCsvDateTime)).TotalMinutes > 30) {
                            DialogResult dialogResult = DialogResult.None;
                            if (dontShowAgain == false) {
                                HandleMissingDataForm handleMissingDataForm = new HandleMissingDataForm(_mainForm);
                                var text = string.Format(
                                            "CSV�t�@�C���Ƀf�[�^�����̉\��������܂�\n" +
                                            "[Excel] {0}�s��: {1}\n" +
                                            "[Excel] {2}�s��: {3}\n\n" +
                                            "�����𑱍s���܂����H",
                                            _excelRownum, prevCsvDateTime, _excelRownum + 1, _csvDateTime);
                                dialogResult = handleMissingDataForm.ShowDialog(text, out dontShowAgain);
                                if (dialogResult == DialogResult.Cancel) {
                                    reader.Close();
                                    _document.Dispose();
                                    File.Delete(_tempFileName!);
                                    return ERROR_RETURN_VALUE;
                                }
                            }
                        }
                        prevCsvDateTime = _csvDateTime;

                        //Excel�̃Z���I�u�W�F�N�g�̍쐬�ƃf�[�^��������
                        _excelRownum++;
                        _mainForm.Invoke((MethodInvoker)(() => {
                            if (_mainForm.progressForm!.StartPos < 0) {
                                _mainForm.progressForm!.StartPos = reader.BaseStream.Position;
                                _mainForm.progressForm!.EndPos = reader.BaseStream.Length;
                            } else {
                                _mainForm.progressForm!.SetProgressBarValue(reader.BaseStream.Position);
                                _mainForm.progressForm!.SetExcelLabelText(_excelRownum);
                            }
                        }));

                        Debug.WriteLine(_excelRownum + "�s: " + _csvDateTime);

                        //�s�ǉ�
                        row = _worksheet!.Descendants<OOXMLS.Row>().FirstOrDefault(r => r.RowIndex! == _excelRownum);
                        if (row == null) {
                            row = new OOXMLS.Row() {
                                RowIndex = Convert.ToUInt32(_excelRownum),
                                //Spans = new ListValue<OOXML.StringValue>()
                            };
                            _sheetData = _worksheet.GetFirstChild<SheetData>();
                            _sheetData!.Append(row);
                        }

                        //1��� ���t
                        date = GetCell(_worksheet, "A" + _excelRownum);
                        if (date == null) {
                            date = new Cell() {
                                CellReference = "A" + _excelRownum,
                                DataType = CellValues.Number,
                                StyleIndex = (UInt32Value)dateStyleIndex,
                            };
                            row.Append(date);
                        }
                        date.CellValue = new OOXMLS.CellValue(System.DateTime.Parse(cols[0]).ToOADate());

                        //2��� ����
                        // ���O�ŃV���A���l�ɕϊ�
                        var hh = Int32.Parse(cols[1].Substring(0, 2));
                        var mm = Int32.Parse(cols[1].Substring(3));
                        var serialValue = (hh + mm / 60.0) / 24.0;
                        time = GetCell(_worksheet, "B" + _excelRownum);
                        if (time == null) {
                            time = new Cell() {
                                CellReference = "B" + _excelRownum,
                                DataType = CellValues.Number,
                                StyleIndex = (UInt32Value)timeStyleIndex,
                            };
                            row.Append(time);
                        }
                        time.CellValue = new OOXMLS.CellValue(serialValue);

                        //3��� ���l
                        wh = GetCell(_worksheet, "C" + _excelRownum);
                        if (wh == null) {
                            wh = new Cell() {
                                CellReference = "C" + _excelRownum,
                                DataType = CellValues.Number,
                            };
                            row.Append(wh);
                        }
                        wh.CellValue = new OOXMLS.CellValue(cols[2]);

                        //4��ځ`�ŏI(11��)�� ���l
                        for (int n = 3; n < cols.Length; n++) {
                            cell = GetCell(_worksheet, ConvertR1C1ToA1(_excelRownum, n + 1));
                            if (cell == null) {
                                cell = new Cell() {
                                    CellReference = ConvertR1C1ToA1(_excelRownum, n + 1),
                                    DataType = CellValues.Number
                                };
                                row.Append(cell);
                            }
                            cell.CellValue = new OOXMLS.CellValue(cols[n]);
                        }
                        _count++;
                    }
                    _excelDateTime = _csvDateTime;
                    _excelWh = Convert.ToDouble(cols[2]);
                    reader.Close();
                    Debug.WriteLine("CSV�t�@�C���ǂݍ��ݏI��");

                    _mainForm.Invoke((MethodInvoker)(() => {
                        _mainForm.WriteExcelButton.BackColor = System.Drawing.Color.LightGreen;
                        _mainForm.ExcelLastDataLabel.Text = string.Format("{0}�s {1} {2}kWh",
                            _excelRownum, _excelDateTime, (int)((_excelWh + 500) / 1000));

                        //30���ȏ��30���A30��������0���A�b��0�b�A�~���b��0�~���b�ɒ����@2�̕��@
                        var dateTime = System.DateTime.Parse(_excelDateTime!);
                        _mainForm.lastDateTime = new System.DateTime(
                            dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour,
                            (dateTime.Minute >= 30) ? 30 : 0, 0, 0);
                    }));
                    Debug.WriteLine("lastDateTime: " + _mainForm.lastDateTime);

                    Debug.WriteLine("ExcelFileUsingOpenXML.Write() ����");
                    return _count;
                });
            }

            public override bool ResizeTable() {
                if (_document == null) {
                    return false;
                }

                //�e�[�u���̐ݒ�
                var tableDefinitionParts = _worksheetPart!.TableDefinitionParts;
                var tableDefinitionPart = tableDefinitionParts.FirstOrDefault(p => p.Table.Name == EXCEL_TABLENAME);
                if (tableDefinitionPart == null) {
                    Debug.WriteLine(EXCEL_TABLENAME + "��������܂���");

                    //�e�[�u���̐V�K�ݒ�
                    if (_worksheetPart.Worksheet.SheetDimension!.Reference == "A1:A1") {
                        _worksheetPart.Worksheet.SheetDimension!.Reference =
                        (OOXML.StringValue)("A1:K" + _excelRownum);
                    }
                    if (AppendTable(_worksheetPart, 1, _excelRownum, 1, 11)
                        == null) {
                        return false;
                    }

                    for (int i = 0; i < (11 - 1 + 1); i++) {
                        Debug.WriteLine(((char)(64 + (i + 1))).ToString() + 1 + " " +
                            GetCellValue(_workbookPart!,
                            GetCell(_worksheet!, ConvertR1C1ToA1(1, i + 1))!));
                    }
                    return true;
                }
                var table = tableDefinitionPart.Table;
                Debug.WriteLine("table.Name: " + table.Name);
                var tableReference = table.Reference;
                var tableLastRownum = GetBottomRownum(tableReference!);
                if (_excelRownum != tableLastRownum) {
                    //�e�[�u���͈͂̍Đݒ�
                    var newTableReference = (OOXML.StringValue)tableReference!.ToString()!.
                        Replace(tableLastRownum.ToString(), _excelRownum.ToString());
                    table.Reference = newTableReference;
                    //�I�[�g�t�B���^�͈͂̍Đݒ�(�e�[�u���ƍ��킹�ĕύX���K�{)
                    table.AutoFilter!.Reference = newTableReference;
                    //SheetDimension�͈͂̍Đݒ�(�e�[�u���ƍ��킹�ĕύX���K�{)
                    _worksheetPart.Worksheet.SheetDimension!.Reference = newTableReference;
                }
                return true;
            }

            public override bool ShowInPane() {
                if (_document == null) {
                    return false;
                }

                //�ŏI�s����ʓ��ɕ\������
                _sheetView = _worksheet!.SheetViews!.GetFirstChild<OOXMLS.SheetView>();
                _pane = _sheetView!.Elements<OOXMLS.Pane>().FirstOrDefault(p => p.ActivePane! == "bottomLeft");
                if (_pane == null) {
                    return false;
                }
                var topLeftCell = (_excelRownum < 8) ? _excelRownum : _excelRownum - 8;
                _pane.TopLeftCell = "A" + topLeftCell;
                return true;
            }

            public override bool ActivateLastCell() {
                if (_document == null) {
                    return false;
                }

                //�ŏI�s��1��ڂ̃Z�����A�N�e�B�u�ɂ���
                _sheetView = _worksheet!.SheetViews!.GetFirstChild<OOXMLS.SheetView>();
                var bottomLeftSelection = _sheetView!.Elements<OOXMLS.Selection>().
                    FirstOrDefault(s => s.Pane! == "bottomLeft");
                if (bottomLeftSelection == null) {
                    return false;
                }
                var cell = "A" + _excelRownum;
                bottomLeftSelection.ActiveCell = cell;
                bottomLeftSelection.SequenceOfReferences =
                    new ListValue<OOXML.StringValue>() {
                        InnerText = cell
                    };
                return true;
            }

            public override bool Close() {
                if (_document == null) {
                    return false;
                }

                //Excel�t�@�C���ۑ�
                if (_count == 0) {
                    _document.Dispose();
                } else {
                    //workbookPart.Workbook.Save(); //���ꂾ�ƕs�\��
                    _document.Save();
                    _document.Close();
                    try {
                        File.Copy(_tempFileName!, EXCEL_FILENAME, true);
                    } catch (Exception) {
                        _mainForm.ShowErrorMessageBox(
                            "�ꎞ�t�@�C����Excel�t�@�C���փR�s�[�ł��܂���ł���!!!\n\n" +
                            "<�Ώ����@> �蓮�ňꎞ�t�@�C��\n" +
                            "(" + _tempFileName + ")��\n" +
                            "Excel�t�@�C���ɏ㏑�����Ă���ꎞ�t�@�C�����폜���Ă�������");
                        Clipboard.SetText(_tempFileName!);
                        return false;
                    }
                }
                File.Delete(_tempFileName!);
                Debug.WriteLine("ExcelFileUsingOpenXML.Close() ����");
                return true;
            }
        }

        private async Task<int> WriteExcelUsingOpenXML(MainForm mainForm) {
            var excelFile = new ExcelFileUsingOpenXML(mainForm);
            DialogResult dialogResult = excelFile.ShowCreateFileDialog(EXCEL_FILENAME);
            if (dialogResult == DialogResult.Cancel) {
                return ERROR_RETURN_VALUE;
            }
            if (dialogResult == DialogResult.OK) {
                if (excelFile.Create(EXCEL_FILENAME) == false) {
                    return ERROR_RETURN_VALUE;
                }
            }
            if (excelFile.Open(EXCEL_FILENAME) == false) {
                return ERROR_RETURN_VALUE;
            }
            var reader = excelFile.OpenCsvFile(CsvFileNameLabel.Text);
            if (reader == null) {
                return ERROR_RETURN_VALUE;
            }
            if (await excelFile.Write(reader) == ERROR_RETURN_VALUE) {
                return ERROR_RETURN_VALUE;
            }
            excelFile.ResizeTable();
            excelFile.ShowInPane();
            excelFile.ActivateLastCell();
            if (excelFile.Close() == false) {
                return ERROR_RETURN_VALUE;
            }
            return excelFile.GetCount();
        }

        private async void WriteExcelButton_Click(object sender, EventArgs e) {
            bool result;

            WriteExcelButton.Enabled = false;
            if (WriteExcelButton.BackColor == System.Drawing.Color.LightGreen) {
                var dialogResult = ShowOKCancelMessageBox(
                    owner: mainForm,
                    text: "Excel�t�@�C���͍ŐV�ł�\n���s���܂����H",
                    button: MessageBoxDefaultButton.Button2);
                if (dialogResult == DialogResult.Cancel) {
                    WriteExcelButton.Enabled = true;
                    return;
                }
            }

            //CSV�t�@�C�����X�g����Ȃ烊�X�g�X�V
            if (UpdateCsvFileListButton.Enabled == false) {
                return;
            }
            if (CsvFileListBox.Items.Count == 0) {
                UpdateCsvFileListButton.Enabled = false;
                progressForm = new ProgressForm(this, "���X�g�X�V");
                result = await csvFileList.Update(progressForm);
                progressForm!.Close();
                UpdateCsvFileListButton.Enabled = true;
                if (result == false) {
                    WriteExcelButton.Enabled = true;
                    return;
                }
            }

            //CSV�t�@�C�����_�E�����[�h
            progressForm = new ProgressForm(this, "CSV�t�@�C���_�E�����[�h");
            result = await flashair.DownloadCSVFile(progressForm);
            progressForm!.Close();
            if (result == false) {
                WriteExcelButton.Enabled = true;
                return;
            }

            WriteExcelButton.Enabled = false;
            progressForm = new ProgressForm(this, "Excel�t�@�C������", ProgressBarStyle.Blocks);
            progressForm.Update();
            int count;
            //Excel�t�@�C���ɏ�������
            count = await WriteExcelUsingOpenXML(this);
            progressForm.Close();
            WriteExcelButton.Enabled = true;

            if (count == ERROR_RETURN_VALUE) {
                return;
            }
            string message;
            if (count == 0) {
                message = "Excel�t�@�C���͍ŐV�ł�";
            } else {
                message = count.ToString() + " �s�������݂܂���";
            }
            Debug.WriteLine(message);
            ShowOKMessageBox(message);
        }

        private void OpenExcelUsingProcessStart(string filePath) {
            var app = new ProcessStartInfo();
            app.FileName = "excel.exe";
            app.Arguments = filePath;
            app.UseShellExecute = true;
            Process.Start(app);
        }

        private void OpenExcelButton_Click(object sender, EventArgs e) {
            //Excel�t�@�C���̑��݂��m�F
            var filePath = Path.GetFullPath(EXCEL_FILENAME);
            Debug.WriteLine("filePath: " + filePath);
            if (File.Exists(filePath) == false) {
                ShowOKMessageBox(
                    EXCEL_FILENAME + "��������܂���ł���", CAPTION_ERROR,
                    MessageBoxIcon.Error);
                return;
            }

            //Excel���N�����ĊJ��
            //Process.Start() ���g�p
            OpenExcelUsingProcessStart(filePath);
        }

        private void timer1_Tick(object sender, EventArgs e) {
            var nowDateTime = System.DateTime.Now;
            //1�b���Ɏ��v�̓������X�V
            ClockLabel.Text = nowDateTime.ToString(CLOCK_FORMAT);

            //30��10�b�ȓ��ł���Ή����������Ȃ�
            if ((nowDateTime - lastDateTime).TotalSeconds <= 1810) {
                return;
            }

            if (WriteExcelButton.BackColor != System.Drawing.Color.Orange) {
                //WriteExcelButton�̃{�^�����I�����W�F��
                WriteExcelButton.BackColor = System.Drawing.Color.Orange;

                //�A�C�R��������Ă�����T�C�Y�𕜌�
                //WindowState�v���p�e�B���g���ăT�C�Y�𕜌�
                if (this.WindowState == FormWindowState.Minimized) {
                    this.WindowState = FormWindowState.Normal;
                }

                //�őO�ʂ�
                this.TopMost = true;
                this.TopMost = false;

                if (progressForm != null && progressForm.IsDisposed == false) {
                    progressForm.TopMost = true;
                    progressForm.TopMost = false;
                }
            }
        }

        private void MainForm_Shown(object sender, EventArgs e) {
            Refresh();
            try {
                var chrormeDriverVersion = new ChromeConfig().GetMatchingBrowserVersion();
                new DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser);
            } catch (AggregateException) {
                ShowErrorMessageBox("�C���^�[�l�b�g�ɐڑ�����Ă��邩�m�F���Ă�������");
                this.Load += (s, e) => Close();
                return;
            } catch (Exception) {
                ShowErrorMessageBox("Chrome�u���E�U�̃o�[�W�������m�F�ł��Ȃ����C���X�g�[������Ă��܂���");
                ChromeRadioButton.Enabled = false;
                EdgeRadioButton.Checked = true;
            }
            try {
                var edgeDriverVersion = new EdgeConfig().GetMatchingBrowserVersion();
                new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);
            } catch (Exception) {
                ShowErrorMessageBox("Edge�u���E�U�̃o�[�W�������m�F�ł��Ȃ����C���X�g�[������Ă��܂���");
                EdgeRadioButton.Enabled = false;
            }
            if (ChromeRadioButton.Enabled == true) {
                Debug.WriteLine("ChromeDriverVersion: " + (new ChromeConfig().GetMatchingBrowserVersion()));
            }
            if (EdgeRadioButton.Enabled == true) {
                Debug.WriteLine("EdgeDriverVersion: " + (new EdgeConfig().GetMatchingBrowserVersion()));
            }
            if ((ChromeRadioButton.Enabled == false) && (EdgeRadioButton.Enabled == false)) {
                ChromeRadioButton.Checked = false;
                EdgeRadioButton.Checked = false;
                UpdateCsvFileListButton.Enabled = false;
                WriteExcelButton.Enabled = false;
            }
        }

        public void ReadBrowserFromInifile() {
            int capacitySize = 256;
            var sb = new StringBuilder(capacitySize);
            var stringLength = GetPrivateProfileString(
                APPNAME, INIFILE_KEY_BROWSER, "", sb, Convert.ToUInt32(sb.Capacity),
                INIFILE_FILENAME);
            if (sb.ToString() == ChromeRadioButton.Text) {
                ChromeRadioButton.Checked = true;
            }
            if (sb.ToString() == EdgeRadioButton.Text) {
                EdgeRadioButton.Checked = true;
            }
        }

        public void WriteBrowserToInifile() {
            if (ChromeRadioButton.Checked == true) {
                WritePrivateProfileString(APPNAME, INIFILE_KEY_BROWSER,
                    ChromeRadioButton.Text, INIFILE_FILENAME);
            }
            if (EdgeRadioButton.Checked == true) {
                WritePrivateProfileString(APPNAME, INIFILE_KEY_BROWSER,
                    EdgeRadioButton.Text, INIFILE_FILENAME);
            }
        }

        private partial class FindFlashairForm : GetFlashairCsv.FindFlashairForm {
            private MainForm _mainForm;
            private FindFlashairForm _findFlashairForm;

            public FindFlashairForm(MainForm mainForm) {
                _mainForm = mainForm;
                _mainForm.FindFlashairButton.Enabled = false;
                _findFlashairForm = this;
                this.CloseButton.Click += CloseButton_Click;
                this.ApplyButton.Click += ApplyButton_Click;
                this.FormClosing += FindFlashairForm_FormClosing;
                //�\���ʒu�̐ݒ�
                var point = _mainForm.Location;
                this.Bounds = new System.Drawing.Rectangle(
                    point.X + 50, point.Y + 80, this.Size.Width, this.Size.Height);
                this.Show();
            }

            private void ApplyButton_Click(object? sender, EventArgs e) {
                //FlashAir��URL�֔��f
                _mainForm.FlashairUrlTextBox.Text = PROTOCOL + _findFlashairForm.IpAddrLabel.Text;
                this.Close();
            }

            private void CloseButton_Click(object? sender, EventArgs e) {
                this.Close();
            }

            private void FindFlashairForm_FormClosing(object? sender, FormClosingEventArgs e) {
                _mainForm.FindFlashairButton.Enabled = true;
            }
        }

        private void FindFlashairButton_Click(object sender, EventArgs e) {
            flashair.ReadMacAddrFromInifile();
            if (flashair.MacAddr == "") {
                ShowErrorMessageBox("FlashAir��MAC�A�h���X�̎w�肪�����ł�");
                return;
            }
            flashair.ReadStartIpAddrFromInifile();
            var startOctet = flashair.StartIpAddr!.Split(".");
            flashair.ReadEndIpAddrFromInifile();
            var endOctet = flashair.EndIpAddr!.Split(".");
            if ((startOctet.Length != 4) || (endOctet.Length != 4)) {
                ShowErrorMessageBox("�����J�nIP�A�h���X�̎w�肪�����ł�");
                return;
            }
            if ((startOctet[0] != endOctet[0]) ||
                (startOctet[1] != endOctet[1]) ||
                (startOctet[2] != endOctet[2])) {
                ShowErrorMessageBox("IP�A�h���X�̌����͈͓͂���Z�O�����g���ł�");
                return;
            }

            //�_�C�A���O�{�b�N�X�\��
            findFlashairForm = new FindFlashairForm(this);
            findFlashairForm.ApplyButton.Enabled = false;

            //IP�A�h���X��������
            //�\�[�X�R�[�h�̌��^���p��
            //[C#] ARP�v���𑗐M���ă����[�gPC��MAC�A�h���X���擾����iSendARP�֐��j
            //https://hensa40.cutegirl.jp/archives/6689/
            //using System.Runtime.InteropServices; ���K�v
            string dstIpAddr; // MAC�A�h���X���擾���郊���[�gPC��IP�A�h���X
            findFlashairForm.FlashairMacAddrLabel.Text = flashair.MacAddr;
            findFlashairForm.StatusLabel.Text = "������...";
            for (int octet = Convert.ToInt32(startOctet[3]);
                octet <= Convert.ToInt32(endOctet[3]); octet++) {
                //�ҋ@���̃C�x���g����������
                Application.DoEvents();

                dstIpAddr = String.Format("{0}.{1}.{2}.{3}",
                    startOctet[0], startOctet[1], startOctet[2], octet.ToString());

                // ������iIP�A�h���X�j����IPAddress�N���X�ɕϊ�
                IPAddress dest = IPAddress.Parse(dstIpAddr);

                // IP�A�h���X�𐔒l�Ƃ��Ď擾����
                // �l�b�g���[�N�o�C�g�I�[�_�ւ̕ϊ��͗v��Ȃ�
                int destAddr = BitConverter.ToInt32(dest.GetAddressBytes(), 0);

                // MAC�A�h���X�p�̃o�b�t�@���m��
                byte[] pMacAddr = new byte[6];
                int PhyAddrLen = pMacAddr.Length;

                // ARP�𑗐M
                findFlashairForm.IpAddrLabel.Text = dstIpAddr;
                int ret;
                try {
                    ret = SendARP(destAddr, 0, pMacAddr, ref PhyAddrLen);
                } catch (Exception ex) {
                    ShowErrorMessageBox(ex);
                    return;
                }
                if (ret == 0) {
                    // ARP�������Ԃ��Ă����ꍇ
                    var dstPhyAddr =
                        String.Format("{0:x2}-{1:x2}-{2:x2}-{3:x2}-{4:x2}-{5:x2}",
                        pMacAddr[0], pMacAddr[1], pMacAddr[2], pMacAddr[3], pMacAddr[4], pMacAddr[5]);
                    Debug.WriteLine(dstIpAddr + " -> " + dstPhyAddr);
                    findFlashairForm.MacAddrLabel.Text = dstPhyAddr;
                    if (dstPhyAddr == flashair.MacAddr) {
                        findFlashairForm.StatusLabel.Text = "FlashAir��������܂���(^_^)";
                        findFlashairForm.IpAddrLabel.ForeColor = System.Drawing.Color.White;
                        findFlashairForm.IpAddrLabel.BackColor = System.Drawing.Color.Green;
                        findFlashairForm.ApplyButton.Enabled = true;
                        findFlashairForm.ApplyButton.Focus();
                        return;
                    }
                } else {
                    // �G���[�R�[�h���o��
                    //const uint FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x100;
                    //const uint FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x2000;
                    //const uint FORMAT_MESSAGE_FROM_HMODULE = 0x800;
                    //const uint FORMAT_MESSAGE_FROM_STRING = 0x400;
                    const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x1000;
                    const uint FORMAT_MESSAGE_IGNORE_INSERTS = 0x200;
                    //const uint FORMAT_MESSAGE_MAX_WIDTH_MASK = 0xFF;
                    const ushort LANG_NEUTRAL = 0;
                    const ushort SUBLANG_DEFAULT = 1;

                    int langId = (SUBLANG_DEFAULT << 10) | LANG_NEUTRAL;
                    int capacitySize = 256;
                    var sb = new StringBuilder(capacitySize);

                    FormatMessage(
                        //FORMAT_MESSAGE_ALLOCATE_BUFFER |  //�e�L�X�g�̃��������蓖�Ă�v������
                        FORMAT_MESSAGE_FROM_SYSTEM |        //�G���[���b�Z�[�W��Windows���p�ӂ��Ă�����̂��g�p
                        FORMAT_MESSAGE_IGNORE_INSERTS,      //���̈����𖳎����ăG���[�R�[�h�ɑ΂���G���[���b�Z�[�W���쐬����
                        0,
                        ret,                                //�G���[�R�[�h
                        langId,                             //������w��
                        sb,                                 //���b�Z�[�W�e�L�X�g���ۑ������o�b�t�@�ւ̃|�C���^
                        Convert.ToUInt32(sb.Capacity),      //�o�b�t�@�̃T�C�Y
                        0);
                    var formattedMessage = sb.ToString().Replace("\r", "").Replace("\n", "");
                    Debug.WriteLine(dstIpAddr + " -> " + formattedMessage);
                    findFlashairForm.MacAddrLabel.Text = formattedMessage;
                }
            }
            findFlashairForm.StatusLabel.Text = "FlashAir��������܂���ł���m(_ _)m";
        }

        private partial class HandleMissingDataForm : GetFlashairCsv.HandleMissingDataForm {
            private MainForm _mainForm;
            private string _text = "";
            private bool _dontShowAgain = false;

            public HandleMissingDataForm(MainForm mainForm) {
                _mainForm = mainForm;
                this.DontShowAgainCheckBox.CheckedChanged += DontShowAgainCheckBox_CheckedChanged!;
                this.OKButton.Click += OKButton_Click!;
                this.CancelButton.Click += CancelButton_Click!;
                this.Load += HandleMissingDataForm_Load!;
            }

            public DialogResult ShowDialog(string text,out bool dontShowAgain) {
                if (text != null) {
                    _text = text;
                }
                var dialogResult = base.ShowDialog();
                dontShowAgain = _dontShowAgain;
                return dialogResult;
            }

            private void HandleMissingDataForm_Load(object sender, EventArgs e) {
                //�\���ʒu�̐ݒ�
                var point = _mainForm.Location;
                this.Bounds = new System.Drawing.Rectangle(
                    point.X + 50, point.Y + 80, this.Size.Width, this.Size.Height);
                this.InformationLabel.Text = _text;
            }

            public void OKButton_Click(object sender, EventArgs e) {
                this.DialogResult = DialogResult.OK;
                this.Dispose();
            }

            public void CancelButton_Click(object sender, EventArgs e) {
                this.DialogResult = DialogResult.Cancel;
                this.Dispose();
            }

            public void DontShowAgainCheckBox_CheckedChanged(object sender, EventArgs e) {
                _dontShowAgain = this.DontShowAgainCheckBox.Checked;
            }
        }
    }
}
