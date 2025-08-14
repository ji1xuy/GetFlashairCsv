using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;
using System.Diagnostics;
using Application = System.Windows.Forms.Application;
using System.Xml;
using System.Net;

//Nuget パッケージのインストール
//Selenium.WebDriver
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;

//Nuget パッケージのインストール
//WebDriverManager
using WebDriverManager;
using WebDriverManager.Helpers;
using WebDriverManager.DriverConfigs.Impl;

//Nuget パッケージのインストール
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
        private const string INIFILE_FILENAME = @"./" + APPNAME + ".ini"; // "./"要
        private const string INIFILE_KEY_URL = "url";
        private const string INIFILE_KEY_BROWSER = "browser";
        private const string PROTOCOL = "http://";
        private const string INIFILE_KEY_MAC_ADDR = "macAddr";
        private const string INIFILE_KEY_START_IP_ADDR = "startIpAddr";
        private const string INIFILE_KEY_END_IP_ADDR = "endIpAddr";
        private const string EXCEL_FILENAME = @"whm_30min.xlsx";
        private const string EXCEL_SHEETNAME = "30分データ";
        private const string EXCEL_TABLENAME = "テーブル1";
        private const string EXCEL_TABLESTYLENAME = "TableStyleMedium16";
        private const string ITEM_NOT_SELECTED = "未選択";
        private const string CAPTION_ERROR = "エラー";
        private const string CAPTION_INFORMATION = "情報";
        private const string CAPTION_QUESTION = "確認";
        private const int ERROR_RETURN_VALUE = -1;
        //EXCEL_HEADER_IDENTIFIER: CSVファイル、Excelファイルのヘッダー(の1列目)の識別子
        private const string EXCEL_HEADER_IDENTIFIER = "yyyy/mm/dd";
        //カスタム日時形式文字列
        //https://learn.microsoft.com/ja-jp/dotnet/standard/base-types/custom-date-and-time-format-strings#H_Specifier
        // Mは月、mは分、hは12時間形式の時間、Hは24時間形式の時間
        // 1桁は先行ゼロなし、2桁は先行ゼロあり)
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
        GetFlashairCsv.iniによる設定方法
        各キーの名前は定数にて設定
        FlashAirのCSVファイルのリスト更新(GUIで設定変更可能)
        ・FlashAirのURL     キー: INIFILE_KEY_URLの値 / 値: http://xxx.xxx.xxx.xxx
        ・使用するブラウザ     キー: INIFILE_KEY_BROWSERの値 / 値: Chrome or Edge
        FlashAirの検索(iniファイルの直接編集のみ)
        ・MACアドレス        キー:　INIFILE_KEY_MAC_ADDRの値 / 値: xx-xx-xx-xx-xx-xx
        ・検索開始IPアドレス   キー:　INIFILE_KEY_START_IP_ADDRの値 / 値: nnn.nnn.nnn.nnn
        ・検索終了IPアドレス   キー:　INIFILE_KEY_END_IP_ADDRの値 / 値: nnn.nnn.nnn.nnn
        (検索開始IPアドレスと検索終了IPアドレスは同一セグメント内であること)
        */

        public MainForm() {
            InitializeComponent();
            mainForm = this;
            Debug.WriteLine("mainForm.Handle: " + mainForm.Handle);
            flashair = new Flashair(this);
            flashair.ReadUrlFromInifile();
            ReadBrowserFromInifile();
            csvFileList = new CsvFileList(this);

            //タイトルバーの設定
            this.Text = WINDOW_TITLE;

            //Excelファイルのフルパスをラベルに表示
            ExcelFileNameLabel.Text =
                Path.Combine(Application.StartupPath, EXCEL_FILENAME);

            WriteExcelButton.BackColor = System.Drawing.Color.Yellow;

            //lastDateTimeの初期値を設定
            var now = System.DateTime.Now;
            //30分以上は30分、30分未満は0分、秒は0秒、ミリ秒は0ミリ秒に調整　2つの方法
            //new DateTime()で新たに作成
            lastDateTime = new System.DateTime(
                now.Year, now.Month, now.Day, now.Hour,
                (now.Minute >= 30) ? 30 : 0, 0, 0);
            Debug.WriteLine("lastDateTime: " + lastDateTime);

            timer1_Tick(Type.Missing, EventArgs.Empty); //引数はダミー
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
                    _mainForm.ShowOKMessageBox("保存しました");
                } else {
                    _mainForm.ShowOKMessageBox("保存できませんでした", CAPTION_ERROR,
                        MessageBoxIcon.Error);
                }
            }

            public void ReadMacAddrFromInifile() {
                //iniファイルから設定を読み込む
                int capacitySize = 256;
                //StringBuilderクラス
                //文字列の追加、置換、挿入を行うと、
                //オブジェクトの内容が変更されるだけで新しいオブジェクトを作成しません
                var sb = new StringBuilder(capacitySize);
                var stringLength = GetPrivateProfileString(
                    APPNAME, INIFILE_KEY_MAC_ADDR, "", sb, Convert.ToUInt32(sb.Capacity),
                    INIFILE_FILENAME);
                //FLashAirのURL をラベルに表示
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
                //CSVファイル名を取得
                if (_mainForm.CsvFileNameLabel.Text == ITEM_NOT_SELECTED) {
                    _mainForm.ShowErrorMessageBox("CSVファイルが選択されていません");
                    return false;
                }

                //FlashAirのURLにCSVファイル名を連結
                var filepath = _mainForm.FlashairUrlTextBox.Text;
                if (filepath.EndsWith("/") == false) {
                    filepath += "/";
                }
                filepath += _mainForm.CsvFileNameLabel.Text;

                //CSVファイルをダウンロード
                var client = new HttpClient();
                //タイムアウト時間の設定（デフォルトは100秒）
                client.Timeout = TimeSpan.FromSeconds(30);
                HttpResponseMessage? response = null;
                try {
                    response = await client.GetAsync(filepath);
                } catch (Exception e) when (e is TaskCanceledException || e is HttpRequestException) {
                    //TaskCanceledException: client.Timeoutで設定したタイムアウトが発生した
                    //HttpRequestException: 接続済みの呼び出し先が一定の時間を過ぎても正しく応答しなかった
                    _mainForm.ShowErrorMessageBox(progressform,
                        "通信中にタイムアウトが発生したためダウンロードを中止しました");
                    return false;
                } catch (Exception e) {
                    _mainForm.ShowErrorMessageBox(progressform, e);
                    return false;
                }
                if (response!.StatusCode != System.Net.HttpStatusCode.OK) {
                    _mainForm.ShowErrorMessageBox(progressform,
                        "CSVファイルをダウンロードできませんでした");
                    return false;
                }
                //保存
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
                        // Webドライバーのインスタンス化
                        ChromeDriverService? chromeService;
                        ChromeOptions chromeOptions = new();
                        chromeOptions.AddArgument("--window-position=" + OUT_OF_SCREEN + "," + OUT_OF_SCREEN);
                        //chromeService = ChromeDriverService.CreateDefaultService();
                        //chromeService = ChromeDriverService.CreateDefaultService(Application.StartupPath);
                        //ドライバの起動場所に自動保存された場所を指定
                        var driverVersion = new ChromeConfig().GetMatchingBrowserVersion();
                        var driverPath = $"./Chrome/{driverVersion}/X64/";
                        chromeService = ChromeDriverService.CreateDefaultService(driverPath);
                        //chromeService.SuppressInitialDiagnosticInformation = true; //診断出力抑制
                        chromeService.HideCommandPromptWindow = true; //コマンドプロンプト画面非表示
                        //chromeOptions.AddArgument("--headless");
                        //Normal: complete(すべてのリソースをダウンロードするのを待ちます)
                        chromeOptions.PageLoadStrategy = PageLoadStrategy.Normal;

                        try {
                            using (driver = new ChromeDriver(chromeService, chromeOptions)) {
                                //Minimize()だとブラウザが一瞬表示されてしまう
                                //driver.Manage().Window.Minimize();
                                //headlessモードではハンドルを取得できない
                                _mainForm.browserHandle = GetBrowserHandle();
                                progressForm.Invoke((MethodInvoker)(() => {
                                    progressForm.AbortButton.Enabled = true;
                                }));
                                driver.Navigate().GoToUrl(_mainForm.flashair.Url!);
                                ReadOnlyCollection<IWebElement> elms =
                                    driver.FindElements(By.XPath(@"//*[@id='thumbnail']/div"));
                                //30分値データのCSVファイル名のリストを取得
                                // ReadOnlyCollection<IWebElement>型のリストから
                                // List<IWebElement>型のリストを生成して
                                // さらに ConvertAll() でList<string>型に変換する
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
                        //ドライバの起動場所に自動保存された場所を指定
                        var driverVersion = new EdgeConfig().GetMatchingBrowserVersion();
                        var driverPath = $"./Edge/{driverVersion}/X64/";
                        edgeService = EdgeDriverService.CreateDefaultService(driverPath);

                        edgeService.HideCommandPromptWindow = true;
                        //上の HideCommandPromptWindow = true でも
                        //コマンドプロンプトが一瞬表示されるため
                        //下の方法(2つ目のAnswer)に変えて無理やり画面外に表示させるようにした
                        //https://stackoverflow.com/questions/35818436/hide-silence-chromedriver-window
                        edgeOptions.AddArgument("--window-position=" + OUT_OF_SCREEN + "," + OUT_OF_SCREEN);
                        //Microsoft Edge WebDriver を入れていなかったのが根本原因
                        //Microsoft Edge WebDriver を入れて
                        //HideCommandPromptWindow = true に戻した

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
                    //リストを空にする
                    _mainForm.Invoke((MethodInvoker)(() => {
                        _mainForm.CsvFileListBox.Items.Clear();
                        _mainForm.CsvFileNameLabel.Text = ITEM_NOT_SELECTED;
                    }));

                    if (list.Count == 0) {
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(progressForm,
                                "CSVファイルが1つも見つかりませんでした\n" +
                                "指定したURLはFlashAirではない可能性があります");
                        }));
                        return false;
                    }

                    //2023/9/1処理追加
                    //ファイル数が多く(20個超過?)なると、
                    //ブラウザではファイル日付毎にグループ化された表示になり
                    //取得したlistの要素にはファイル日付が混在するようになる
                    //(例) "20230901\r\n202309.CSV"
                    //ファイル名とファイル日付を別要素に分割しておく
                    var s = String.Join("\r\n", list);
                    list = s.Split("\r\n", StringSplitOptions.None).ToList();

                    //list を条件で絞り込み
                    //RemoveAll() とラムダ式で処理
                    list.RemoveAll(s => !Regex.IsMatch(s, @"[0-9]{6}.CSV"));

                    //list を降順ソート
                    //LINQで降順ソート
                    list.OrderByDescending(s => s);

                    //リストボックスに追加
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
                Debug.WriteLine("GetBrowserHandle()の戻り値: " + browserHandle);
                return browserHandle;
            }

            private void HandleException(ProgressForm progressForm, Exception e) {
                switch (e) {
                    case InvalidOperationException:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(
                                progressForm,
                                "ブラウザを操作できませんでした\n" +
                                "Selenium.WebDriverとブラウザのバージョンが一致しているか確認してください");
                        }));
                        break;
                    case NoSuchWindowException:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(
                                progressForm, "処理を中断しました");
                        }));
                        break;
                    //case WebDriverArgumentException:
                    case WebDriverException:
                        _mainForm.Invoke((MethodInvoker)(() => {
                            _mainForm.ShowErrorMessageBox(
                                progressForm,
                                "FlashAirと通信できませんでした\n" +
                                "FlashAirのURLが正しいか確認してください");
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

        //partial修飾子を付けてProgressForm.csに書かずここに記述
        private partial class ProgressForm : GetFlashairCsv.ProgressForm {
            MainForm _mainForm;
            string _caption;

            public ProgressForm(MainForm mainForm, string caption, ProgressBarStyle progressBarStyle = ProgressBarStyle.Marquee) {
                var systemMenu = GetSystemMenu(this.Handle, false);
                _mainForm = mainForm;
                _caption = caption;

                RemoveMenu(systemMenu, 5, MF_BYPOSITION);
                //閉じるボタンを無効化
                RemoveMenu(systemMenu, SC_CLOSE, MF_BYCOMMAND);
                this.AbortButton.Click += AbortButton_Click;
                this.FormClosing += ProgressForm_FormClosing;
                //表示位置の設定
                var point = _mainForm.Location;
                this.Bounds = new System.Drawing.Rectangle(
                    point.X + 100, point.Y + 150, this.Size.Width, this.Size.Height);
                this.Text = caption;
                this.ProgressBar.Style = progressBarStyle;
                //処理中フォームを表示
                //.Show()でモードレス、.ShowDialog()でモーダル
                this.Show();
                //progressForm のラベルをすぐ表示する
                //無効領域(画面更新が必要な領域)を再描画する
                this.Update();
            }

            private void ProgressForm_FormClosing(object? sender, FormClosingEventArgs e) {
                //ブラウザが"FlashAir"のタイトルのまま残る場合があるため終了させる
                if (_mainForm.browserHandle != null) {
                    SendMessage((IntPtr)_mainForm.browserHandle, WM_CLOSE, 0, 0);
                }
            }

            private void AbortButton_Click(object? sender, EventArgs e) {
                //ブラウザを終了しGoToUrl()処理でエラーを発生させることで中断する
                Debug.WriteLine("GoToUrl()処理中断中...");
                this.StatusLabel.Text = "中断中...";
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
                    Debug.WriteLine("SendMessage()実行");
                    SendMessage((IntPtr)_mainForm.browserHandle, WM_CLOSE, 0, 0);
                }
            }

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public long StartPos { get; set; } = -1;

            [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
            public long EndPos { private get; set; }

            public void SetExcelLabelText(long row) {
                this.ExcelLabel.Text = "[Excel] " + row + "行目書込";
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
                this.CsvLabel.Text = "[CSV] " + EndPos + "バイト中、" + position + "バイト読込";
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
            //処理中のフォームを表示
            progressForm = new ProgressForm(this, "リスト更新");

            await csvFileList.Update(progressForm);

            //処理中のフォームを閉じる
            progressForm.Close();
            UpdateCsvFileListButton.Enabled = true;
        }

        private void CloseButton_Click(object sender, EventArgs e) {
            if (browserHandle != null) {
                Debug.WriteLine("SendMessage()実行");
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

            //Excelファイルの存在を確認
            //見つかった場合は戻り値Noneを返す
            //見つからない場合、新規作成するならOK、作成しないならCancelを返す
            protected DialogResult ShowCreateFileDialog(MainForm form, string fileName) {
                //Excelファイルの存在を確認
                var filePath = Path.GetFullPath(fileName);
                Debug.WriteLine("filePath: " + filePath);
                if (File.Exists(filePath)) {
                    return DialogResult.None;
                }
                //Excelファイルが見つからない場合
                //新規作成するかどうか確認
                return form.ShowOKCancelMessageBox(
                    owner: form,
                    text: EXCEL_FILENAME + " が見つかりません\n新規作成しますか？");
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
            //MainFormのコントロール、フィールドの読み書き
            //メソッドの実行のための変数
            //このコンストラクターの引数からセットされる
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

            //コンストラクター
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
            /// 列番号からエクセルの列名を得る　(例 5 → "E"）
            /// </summary>
            /// <param name="c"></param>
            /// <returns></returns>
            public string ConvertR1C1ToA1(int r, int c) {
                //A～Z列までならこの1行でOK
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
                //階層構造　Worksheet -> (子)SheetData -> (孫)Row -> (曾孫)Cell
                //worksheet.Elements<cell>()では子要素しか取得できない
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

                //テーブルのセル範囲
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

                //worksheetPartからworkbookPartを取得
                var workbookPart = (WorkbookPart?)worksheetPart.GetParentParts().FirstOrDefault(p => p is WorkbookPart);
                for (int i = 0; i < (colMax - colMin + 1); i++) {
                    tableColumns.Append(new TableColumn() {
                        Id = (UInt32)(colMin + i),
                        //Nameは、設定対象のセルに格納されている値と同じ内容を設定する
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
                // is not null を != null にするとエラー
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
                        //runProperties.Append(new RunFont() { Val = "MS P ゴシック" });
                        runProperties.Append(new RunFont() { Val = "ＭＳ Ｐゴシック" });
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

                    //シートの作成
                    var sheetData = new SheetData();

                    //A列の作成
                    var row = new OOXMLS.Row() {
                        RowIndex = 1U,
                        Spans = new ListValue<OOXML.StringValue>(),
                        Height = 36U,
                        CustomHeight = true
                    };

                    //A1セルの作成
                    Cell cell;
                    cell = new Cell() {
                        CellReference = "A1",
                        //R1C1形式でも可。ただし参照の際は作成した形式でしか参照できない
                        //Excelで一度保存し直せばA1形式に変換される様子
                        DataType = CellValues.String,
                        CellValue = new OOXMLS.CellValue(EXCEL_HEADER_IDENTIFIER),
                    };
                    row.Append(cell);
                    sheetData.Append(row);

                    var worksheet = new OOXMLS.Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
                    //worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    //worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                    worksheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

                    //データ範囲の設定
                    var sheetDimension =
                        new SheetDimension() { Reference = "A1:A1" };

                    //スタイルシートの追加
                    var stylesPart =
                        workbookPart.AddNewPart<WorkbookStylesPart>();
                    var stylesheet = new OOXMLS.Stylesheet();

                    //フォント定義
                    var fonts = new OOXMLS.Fonts() { Count = 1 };
                    var font = new OOXMLS.Font() {
                        FontSize = new OOXMLS.FontSize() { Val = 11 },
                        FontName = new FontName() { Val = new StringValue("游ゴシック") },
                        FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                        FontCharSet = new OOXMLS.FontCharSet() { Val = 128 },
                        FontScheme = new OOXMLS.FontScheme() { Val = FontSchemeValues.Minor }
                    };
                    fonts.Append(font);
                    stylesheet.Append(fonts);

                    //塗りつぶしの定義
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

                    //ボーダー定義
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

                    //セルスタイルフォーマット定義
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

                    //セルフォーマット定義を追加
                    var cellFormats = new OOXMLS.CellFormats();
                    //標準用
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
                    //文字折り返し用
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

                    //シートビューの作成
                    var sheetViews = new OOXMLS.SheetViews();
                    SheetView sheetView = new SheetView() {
                        TabSelected = true,
                        WorkbookViewId = (UInt32Value)0U
                    };

                    //ウィンドウ枠の固定の設定
                    //selection1より後にAppendするとエラーになる
                    var pane = new OOXMLS.Pane() {
                        //HorizontalSplit = 1, //列固定
                        VerticalSplit = 1,     //行固定
                        TopLeftCell = "A2",    //下のペインの左上のセル
                        ActivePane = PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen
                    };
                    sheetView.Append(pane);

                    //アクティブセルの設定
                    var selection = new OOXMLS.Selection() {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = "A1",
                        SequenceOfReferences = new ListValue<OOXML.StringValue>() {
                            InnerText = "A1" //R1C1形式は不可
                        }
                    };
                    sheetView.Append(selection);
                    sheetViews.Append(sheetView);

                    worksheet.Append(sheetDimension);
                    worksheet.Append(sheetViews);

                    //ページ設定
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

                    //列幅の設定
                    //sheetData1より後にAppendするとエラーになる
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

                    //コメントの挿入
                    //Excelで開いてコメントの編集状態で名前ボックスに表示される
                    //名前(長い16進数が設定されている)を設定する方法が見つからなかった
                    //そもそもExcelでコメントの挿入をしても名前の変更はできない
                    List<string> ColumnName = new List<string>() { "A" };
                    List<string> CellIndex = new List<string>() { "1" };
                    List<string> NewCommentList = new List<string>() { "シート識別のため変更不可" };
                    InsertComments(worksheetPart, ColumnName, CellIndex, NewCommentList);

                    _document.Save();
                    _document.Close();

                    Debug.WriteLine("ExcelFileUsingOpenXML.Create() 成功");
                    return true;
                } catch (Exception e) {
                    Debug.WriteLine(e);
                    _mainForm.ShowErrorMessageBox(fileName + " を新規作成できませんでした");
                    return false;
                }
            }

            //Open()について
            //Open()したらDispose()しないとファイルが破損する
            //ただしファイル(のタイムスタンプ？)が更新されてしまう
            //一切保存せずオブジェクトを破棄する方法がない？
            //↓回避策　
            //Open()にて処理　ファイルを開く時
            //一時ファイルを作り一時ファイル上で更新する
            //Close()にて処理　保存時
            //元のファイルに上書きして一時ファイルを削除する
            public override bool Open(string fileName) {
                if (String.IsNullOrEmpty(fileName)) {
                    return false;
                }

                try {
                    try {
                        _tempFileName = Path.GetTempFileName();
                        Debug.WriteLine("tempFileName: " + _tempFileName);
                        //tempFileName: C:\Users\(ユーザー名)\AppData\Local\Temp\tmp????.tmp
                    } catch (Exception) {
                        _mainForm.ShowErrorMessageBox("一時ファイルを作成できませんでした");
                        return false;
                    }
                    try {
                        File.Copy(EXCEL_FILENAME, _tempFileName, true);
                    } catch {
                        _mainForm.ShowErrorMessageBox(
                            "Excelファイルを一時ファイルへコピーできませんでした");
                        return false;
                    }
                    _document = SpreadsheetDocument.Open(_tempFileName, true,
                       new OpenSettings { AutoSave = false });
                } catch (System.IO.IOException) {
                    _mainForm.ShowErrorMessageBox(EXCEL_FILENAME + "にアクセスできません");
                    return false;
                } catch (Exception e) {
                    _mainForm.ShowErrorMessageBox(e);
                    return false;
                }

                //Excelファイルの中身が正しいか確認
                _workbookPart = _document.WorkbookPart;
                var sheets =
                    _workbookPart!.Workbook.GetFirstChild<OOXMLS.Sheets>();
                _sheet =
                    sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == EXCEL_SHEETNAME);
                if (_sheet == null) {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox(EXCEL_SHEETNAME + "が見つかりません");
                    return false;
                }
                Debug.WriteLine("sheet.Name: " + _sheet.Name);
                _worksheetPart = (WorksheetPart)_workbookPart.GetPartById(_sheet.Id!);
                _worksheet = _worksheetPart.Worksheet;
                var a1 = GetCell(_worksheet, "A1");
                if (a1 == null) {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox("A1セルが見つかりません");
                    return false;
                }
                var cellValue = GetCellValue(_workbookPart, a1);
                Debug.WriteLine("A1: " + cellValue);
                if (cellValue != EXCEL_HEADER_IDENTIFIER) {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox(
                        "30分値データが見つかりません\n" +
                        "A1セル≠" + EXCEL_HEADER_IDENTIFIER);
                    return false;
                }

                //データのセル範囲を取得
                var usedRange = _worksheetPart.Worksheet.SheetDimension!.Reference;
                Debug.WriteLine("usedRange: " + usedRange);

                //最終行を取得
                _excelRownum = GetBottomRownum(usedRange!);
                Debug.WriteLine("excelRownum: " + _excelRownum);

                //最終行のデータを取得
                if (_excelRownum == 1) {
                    //ヘッダー行しかない場合
                    _excelDateTime = "";
                    _excelWh = 0.0;
                } else {
                    var date = GetCell(_worksheet, "A" + _excelRownum);
                    var time = GetCell(_worksheet, "B" + _excelRownum);
                    var wh = GetCell(_worksheet, "C" + _excelRownum);
                    try {
                        _excelDateTime = "";
                        _excelWh = 0.0;
                        //手動でヘッダー行以外を全削除してある場合の対処
                        //テーブルが設定されていると見た目はヘッダー行しかなくても
                        //2行目にRowオブジェクトとCellオブジェクト(InnerTextが"")が存在し
                        //_excelRownumの値が2と認識される
                        if (date!.InnerText == "") {
                            _excelRownum = 1;
                        } else {
                            var dateSerial = double.Parse(date.InnerText);
                            var timeSerial = double.Parse(time!.InnerText);
                            _excelDateTime = ConvertDateTimeToString(
                                System.DateTime.FromOADate(dateSerial),
                                System.DateTime.FromOADate(timeSerial));
                            // 月末の最後のデータまで書込がされていないのに次の月を書込みしようとした場合に確認
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
                                               "{0}年{1}月分が最後まで書き込みされていません\n\n" +
                                               "書き込みを続行しますか？", dateTime.Year, dateTime.Month),
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
                    _mainForm.ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                        _excelRownum, _excelDateTime, (int)((_excelWh + 500) / 1000));
                }
                Debug.WriteLine("ExcelFileUsingOpenXML.Open() 成功");
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
                    _mainForm.ShowErrorMessageBox("CSVファイルが見つかりません");
                    return null;
                } catch (Exception e) {
                    _document.Dispose();
                    _mainForm.ShowErrorMessageBox(e);
                    return null;
                }

                //CSVファイルのヘッダーを読み込み
                string line = reader.ReadLine()!;
                if (line == null) {
                    _document.Dispose();
                    _mainForm.ShowErrorMessageBox("CSVファイルの中身が空です");
                    return null;
                }
                string[] cols = line.Split(',');
                //Excelファイルのヘッダーが仮ヘッダーだったら
                var b1 = GetCell(_worksheet!, "B1");
                if (b1 == null) {
                    //CSVファイルのヘッダーをコピー
                    //ヘッダーのセル書式
                    //セル書式が追加済か確認
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
                Debug.WriteLine("ExcelFileUsingOpenXML.OpenCsvFile() 成功");
                return reader;
            }

            public override async Task<int> Write(StreamReader reader) {
                return await Task.Run(() => {
                    if ((_document == null) || (reader == null)) {
                        return ERROR_RETURN_VALUE;
                    }
                    //Excelスタイルシートにセルフォーマット定義を追加
                    //Cellの書式を設定する(OpenXML編)
                    //https://www.cloverfield.co.jp/2020/02/28/cell%E3%81%AE%E6%9B%B8%E5%BC%8F%E3%82%92%E8%A8%AD%E5%AE%9A%E3%81%99%E3%82%8Bopenxml%E7%B7%A8/
                    var stylesPart = _workbookPart!.WorkbookStylesPart;
                    var cellFormats = stylesPart!.Stylesheet.CellFormats;

                    //日付のセル書式
                    ////NumberFormatId=14の組込み書式 [日付]種類 *2012/3/14
                    //セル書式が追加済か確認
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
                    //セル書式を追加
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

                    //時刻のセル書式
                    //Apply OpenXML excel number formatcode to string value in C#
                    //https://stackoverflow.com/questions/25228471/apply-openxml-excel-number-formatcode-to-string-value-in-c-sharp
                    //https://learn.microsoft.com/ja-jp/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-2.8.1
                    //NumberFormatId=20の組込み書式"h:mm"は、[時刻]種類13:30にできない
                    //FormatCode="h:mm;@"のユーザー定義を追加する必要がある
                    //ユーザー定義のNumberFormatIdは176以降らしい
                    //176を初期値にして順番で見ていって空いていたら発番することにした
                    //エラーは出ていないが、この方法で全く問題ないのか不明
                    //<NumberFormatIdの発番の仕組みを検証>
                    //セルに対して組込みの[時刻]種類13:30をExcelで操作して設定すると
                    //[時刻]種類13:30になるが、styles.xmlを覗くとユーザー定義になり
                    //FormatCodeは"h:mm;@"で,NumberFormatIdの発番は
                    //176は空いているのに176にならず、やる度に180,181と一定しない
                    // ファイルではなくExcelがグローバルに管理している?(PCが別なら管理は無理)
                    // ランダムに番号を作成してファイル内で未使用なら決定?
                    // 2つのファイル(同じIdでCodeの内容が違う)を1つにした時どうなるのか?

                    // "h:mm;@"のユーザー定義が登録済か確認
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

                    //セル書式が追加済か確認
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

                    //セル書式を追加
                    if (timeStyleIndex! == null) {
                        //セル書式を作成
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

                    //Excelファイルに書き込み
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
                        //時間がかかる処理での「応答なし」を回避するには？
                        //https://atmarkit.itmedia.co.jp/ait/articles/0403/19/news088.html
                        //Application.DoEvents();

                        // 読み込んだ文字列から制御文字を除去
                        // 空の行なら処理をスキップ
                        line = string.Concat(reader.ReadLine()!.Where(c => !char.IsControl(c)));
                        if (line == "") {
                            continue;
                        }
                        // 文字列をカンマ区切りで配列に格納
                        cols = line.Split(',');
                        _csvDateTime = cols[0] + " " + cols[1];
                        //_csvDateTimeの書式: yyyy/mm/dd hh:mm
                        //Excel最終行の日時より日時が同じか前ならスキップ
                        if ((System.DateTime.Parse(_csvDateTime) - System.DateTime.Parse(_excelDateTime!)).TotalMinutes <= 0) {
                            continue;
                        }
                        //前回のループで処理したデータの日時との差が30分より大きければ確認
                        if ((System.DateTime.Parse(_csvDateTime) - System.DateTime.Parse(prevCsvDateTime)).TotalMinutes > 30) {
                            DialogResult dialogResult = DialogResult.None;
                            if (dontShowAgain == false) {
                                HandleMissingDataForm handleMissingDataForm = new HandleMissingDataForm(_mainForm);
                                var text = string.Format(
                                            "CSVファイルにデータ欠落の可能性があります\n" +
                                            "[Excel] {0}行目: {1}\n" +
                                            "[Excel] {2}行目: {3}\n\n" +
                                            "書込を続行しますか？",
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

                        //Excelのセルオブジェクトの作成とデータ書き込み
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

                        Debug.WriteLine(_excelRownum + "行: " + _csvDateTime);

                        //行追加
                        row = _worksheet!.Descendants<OOXMLS.Row>().FirstOrDefault(r => r.RowIndex! == _excelRownum);
                        if (row == null) {
                            row = new OOXMLS.Row() {
                                RowIndex = Convert.ToUInt32(_excelRownum),
                                //Spans = new ListValue<OOXML.StringValue>()
                            };
                            _sheetData = _worksheet.GetFirstChild<SheetData>();
                            _sheetData!.Append(row);
                        }

                        //1列目 日付
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

                        //2列目 時刻
                        // 自前でシリアル値に変換
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

                        //3列目 数値
                        wh = GetCell(_worksheet, "C" + _excelRownum);
                        if (wh == null) {
                            wh = new Cell() {
                                CellReference = "C" + _excelRownum,
                                DataType = CellValues.Number,
                            };
                            row.Append(wh);
                        }
                        wh.CellValue = new OOXMLS.CellValue(cols[2]);

                        //4列目～最終(11列)目 数値
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
                    Debug.WriteLine("CSVファイル読み込み終了");

                    _mainForm.Invoke((MethodInvoker)(() => {
                        _mainForm.WriteExcelButton.BackColor = System.Drawing.Color.LightGreen;
                        _mainForm.ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                            _excelRownum, _excelDateTime, (int)((_excelWh + 500) / 1000));

                        //30分以上は30分、30分未満は0分、秒は0秒、ミリ秒は0ミリ秒に調整　2つの方法
                        var dateTime = System.DateTime.Parse(_excelDateTime!);
                        _mainForm.lastDateTime = new System.DateTime(
                            dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour,
                            (dateTime.Minute >= 30) ? 30 : 0, 0, 0);
                    }));
                    Debug.WriteLine("lastDateTime: " + _mainForm.lastDateTime);

                    Debug.WriteLine("ExcelFileUsingOpenXML.Write() 成功");
                    return _count;
                });
            }

            public override bool ResizeTable() {
                if (_document == null) {
                    return false;
                }

                //テーブルの設定
                var tableDefinitionParts = _worksheetPart!.TableDefinitionParts;
                var tableDefinitionPart = tableDefinitionParts.FirstOrDefault(p => p.Table.Name == EXCEL_TABLENAME);
                if (tableDefinitionPart == null) {
                    Debug.WriteLine(EXCEL_TABLENAME + "が見つかりません");

                    //テーブルの新規設定
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
                    //テーブル範囲の再設定
                    var newTableReference = (OOXML.StringValue)tableReference!.ToString()!.
                        Replace(tableLastRownum.ToString(), _excelRownum.ToString());
                    table.Reference = newTableReference;
                    //オートフィルタ範囲の再設定(テーブルと合わせて変更が必須)
                    table.AutoFilter!.Reference = newTableReference;
                    //SheetDimension範囲の再設定(テーブルと合わせて変更が必須)
                    _worksheetPart.Worksheet.SheetDimension!.Reference = newTableReference;
                }
                return true;
            }

            public override bool ShowInPane() {
                if (_document == null) {
                    return false;
                }

                //最終行を画面内に表示する
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

                //最終行の1列目のセルをアクティブにする
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

                //Excelファイル保存
                if (_count == 0) {
                    _document.Dispose();
                } else {
                    //workbookPart.Workbook.Save(); //これだと不十分
                    _document.Save();
                    _document.Close();
                    try {
                        File.Copy(_tempFileName!, EXCEL_FILENAME, true);
                    } catch (Exception) {
                        _mainForm.ShowErrorMessageBox(
                            "一時ファイルをExcelファイルへコピーできませんでした!!!\n\n" +
                            "<対処方法> 手動で一時ファイル\n" +
                            "(" + _tempFileName + ")を\n" +
                            "Excelファイルに上書きしてから一時ファイルを削除してください");
                        Clipboard.SetText(_tempFileName!);
                        return false;
                    }
                }
                File.Delete(_tempFileName!);
                Debug.WriteLine("ExcelFileUsingOpenXML.Close() 成功");
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
                    text: "Excelファイルは最新です\n実行しますか？",
                    button: MessageBoxDefaultButton.Button2);
                if (dialogResult == DialogResult.Cancel) {
                    WriteExcelButton.Enabled = true;
                    return;
                }
            }

            //CSVファイルリストが空ならリスト更新
            if (UpdateCsvFileListButton.Enabled == false) {
                return;
            }
            if (CsvFileListBox.Items.Count == 0) {
                UpdateCsvFileListButton.Enabled = false;
                progressForm = new ProgressForm(this, "リスト更新");
                result = await csvFileList.Update(progressForm);
                progressForm!.Close();
                UpdateCsvFileListButton.Enabled = true;
                if (result == false) {
                    WriteExcelButton.Enabled = true;
                    return;
                }
            }

            //CSVファイルをダウンロード
            progressForm = new ProgressForm(this, "CSVファイルダウンロード");
            result = await flashair.DownloadCSVFile(progressForm);
            progressForm!.Close();
            if (result == false) {
                WriteExcelButton.Enabled = true;
                return;
            }

            WriteExcelButton.Enabled = false;
            progressForm = new ProgressForm(this, "Excelファイル書込", ProgressBarStyle.Blocks);
            progressForm.Update();
            int count;
            //Excelファイルに書き込む
            count = await WriteExcelUsingOpenXML(this);
            progressForm.Close();
            WriteExcelButton.Enabled = true;

            if (count == ERROR_RETURN_VALUE) {
                return;
            }
            string message;
            if (count == 0) {
                message = "Excelファイルは最新です";
            } else {
                message = count.ToString() + " 行書き込みました";
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
            //Excelファイルの存在を確認
            var filePath = Path.GetFullPath(EXCEL_FILENAME);
            Debug.WriteLine("filePath: " + filePath);
            if (File.Exists(filePath) == false) {
                ShowOKMessageBox(
                    EXCEL_FILENAME + "が見つかりませんでした", CAPTION_ERROR,
                    MessageBoxIcon.Error);
                return;
            }

            //Excelを起動して開く
            //Process.Start() を使用
            OpenExcelUsingProcessStart(filePath);
        }

        private void timer1_Tick(object sender, EventArgs e) {
            var nowDateTime = System.DateTime.Now;
            //1秒毎に時計の日時を更新
            ClockLabel.Text = nowDateTime.ToString(CLOCK_FORMAT);

            //30分10秒以内であれば何も処理しない
            if ((nowDateTime - lastDateTime).TotalSeconds <= 1810) {
                return;
            }

            if (WriteExcelButton.BackColor != System.Drawing.Color.Orange) {
                //WriteExcelButtonのボタンをオレンジ色に
                WriteExcelButton.BackColor = System.Drawing.Color.Orange;

                //アイコン化されていたらサイズを復元
                //WindowStateプロパティを使ってサイズを復元
                if (this.WindowState == FormWindowState.Minimized) {
                    this.WindowState = FormWindowState.Normal;
                }

                //最前面に
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
                ShowErrorMessageBox("インターネットに接続されているか確認してください");
                this.Load += (s, e) => Close();
                return;
            } catch (Exception) {
                ShowErrorMessageBox("Chromeブラウザのバージョンが確認できないかインストールされていません");
                ChromeRadioButton.Enabled = false;
                EdgeRadioButton.Checked = true;
            }
            try {
                var edgeDriverVersion = new EdgeConfig().GetMatchingBrowserVersion();
                new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);
            } catch (Exception) {
                ShowErrorMessageBox("Edgeブラウザのバージョンが確認できないかインストールされていません");
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
                //表示位置の設定
                var point = _mainForm.Location;
                this.Bounds = new System.Drawing.Rectangle(
                    point.X + 50, point.Y + 80, this.Size.Width, this.Size.Height);
                this.Show();
            }

            private void ApplyButton_Click(object? sender, EventArgs e) {
                //FlashAirのURLへ反映
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
                ShowErrorMessageBox("FlashAirのMACアドレスの指定が無効です");
                return;
            }
            flashair.ReadStartIpAddrFromInifile();
            var startOctet = flashair.StartIpAddr!.Split(".");
            flashair.ReadEndIpAddrFromInifile();
            var endOctet = flashair.EndIpAddr!.Split(".");
            if ((startOctet.Length != 4) || (endOctet.Length != 4)) {
                ShowErrorMessageBox("検索開始IPアドレスの指定が無効です");
                return;
            }
            if ((startOctet[0] != endOctet[0]) ||
                (startOctet[1] != endOctet[1]) ||
                (startOctet[2] != endOctet[2])) {
                ShowErrorMessageBox("IPアドレスの検索範囲は同一セグメント内です");
                return;
            }

            //ダイアログボックス表示
            findFlashairForm = new FindFlashairForm(this);
            findFlashairForm.ApplyButton.Enabled = false;

            //IPアドレス検索処理
            //ソースコードの原型引用元
            //[C#] ARP要求を送信してリモートPCのMACアドレスを取得する（SendARP関数）
            //https://hensa40.cutegirl.jp/archives/6689/
            //using System.Runtime.InteropServices; が必要
            string dstIpAddr; // MACアドレスを取得するリモートPCのIPアドレス
            findFlashairForm.FlashairMacAddrLabel.Text = flashair.MacAddr;
            findFlashairForm.StatusLabel.Text = "検索中...";
            for (int octet = Convert.ToInt32(startOctet[3]);
                octet <= Convert.ToInt32(endOctet[3]); octet++) {
                //待機中のイベントを処理する
                Application.DoEvents();

                dstIpAddr = String.Format("{0}.{1}.{2}.{3}",
                    startOctet[0], startOctet[1], startOctet[2], octet.ToString());

                // 文字列（IPアドレス）からIPAddressクラスに変換
                IPAddress dest = IPAddress.Parse(dstIpAddr);

                // IPアドレスを数値として取得する
                // ネットワークバイトオーダへの変換は要らない
                int destAddr = BitConverter.ToInt32(dest.GetAddressBytes(), 0);

                // MACアドレス用のバッファを確保
                byte[] pMacAddr = new byte[6];
                int PhyAddrLen = pMacAddr.Length;

                // ARPを送信
                findFlashairForm.IpAddrLabel.Text = dstIpAddr;
                int ret;
                try {
                    ret = SendARP(destAddr, 0, pMacAddr, ref PhyAddrLen);
                } catch (Exception ex) {
                    ShowErrorMessageBox(ex);
                    return;
                }
                if (ret == 0) {
                    // ARP応答が返ってきた場合
                    var dstPhyAddr =
                        String.Format("{0:x2}-{1:x2}-{2:x2}-{3:x2}-{4:x2}-{5:x2}",
                        pMacAddr[0], pMacAddr[1], pMacAddr[2], pMacAddr[3], pMacAddr[4], pMacAddr[5]);
                    Debug.WriteLine(dstIpAddr + " -> " + dstPhyAddr);
                    findFlashairForm.MacAddrLabel.Text = dstPhyAddr;
                    if (dstPhyAddr == flashair.MacAddr) {
                        findFlashairForm.StatusLabel.Text = "FlashAirが見つかりました(^_^)";
                        findFlashairForm.IpAddrLabel.ForeColor = System.Drawing.Color.White;
                        findFlashairForm.IpAddrLabel.BackColor = System.Drawing.Color.Green;
                        findFlashairForm.ApplyButton.Enabled = true;
                        findFlashairForm.ApplyButton.Focus();
                        return;
                    }
                } else {
                    // エラーコードを出力
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
                        //FORMAT_MESSAGE_ALLOCATE_BUFFER |  //テキストのメモリ割り当てを要求する
                        FORMAT_MESSAGE_FROM_SYSTEM |        //エラーメッセージはWindowsが用意しているものを使用
                        FORMAT_MESSAGE_IGNORE_INSERTS,      //次の引数を無視してエラーコードに対するエラーメッセージを作成する
                        0,
                        ret,                                //エラーコード
                        langId,                             //言語を指定
                        sb,                                 //メッセージテキストが保存されるバッファへのポインタ
                        Convert.ToUInt32(sb.Capacity),      //バッファのサイズ
                        0);
                    var formattedMessage = sb.ToString().Replace("\r", "").Replace("\n", "");
                    Debug.WriteLine(dstIpAddr + " -> " + formattedMessage);
                    findFlashairForm.MacAddrLabel.Text = formattedMessage;
                }
            }
            findFlashairForm.StatusLabel.Text = "FlashAirが見つかりませんでしたm(_ _)m";
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
                //表示位置の設定
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
