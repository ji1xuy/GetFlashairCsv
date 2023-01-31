using System.Runtime.InteropServices;
using System.Text;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Application = System.Windows.Forms.Application;
using System.Xml;

//Nuget パッケージのインストール
//DotNetCore.NPOI
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

//Nuget パッケージのインストール
//Selenium.WebDriver
//Selenium.WebDriver.ChromeDriver
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;

//Nuget パッケージのインストール
//ClosedXML
using ClosedXML.Excel;

//Nuget パッケージのインストール
//DocumentFormat.OpenXml
using DocumentFormat.OpenXml;
using OOXML = DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OOXMLS = DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

//COM参照の追加
//Microsoft Excel 16.0 ObjectLibrary
using Microsoft.Office.Interop.Excel;
using ExcelApplication = Microsoft.Office.Interop.Excel.Application;

//Microsoft Edge WebDriverのダウンロード
//https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/
//Stable チャネル バージョン:  108.0.1462.54: x64

/*

*/

namespace GetFlashairCsv
{
    public partial class MainForm : Form
    {
        private const string APPNAME = "GetFlashairCsv";
        private const string WINDOW_TITLE = APPNAME + "_20230129";
        private const string INI_FILENAME = @"./" + APPNAME + ".ini"; // "./"要
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
        //private static MainForm mainForm;
        private MainForm mainForm;
        private Flashair flashair;
        private CsvFileList csvFileList;
        private ProgressForm? progressForm;
        private DateTime lastDateTime;

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

        public MainForm()
        {
            InitializeComponent();
            mainForm = this;
            Debug.WriteLine("mainForm.Handle: " + mainForm.Handle);
            flashair = new Flashair(this);
            flashair.ReadUrlFromInifile();
            csvFileList = new CsvFileList(this);

            //タイトルバーの設定
            this.Text = WINDOW_TITLE;

            //Excelファイルのフルパスをラベルに表示
            ExcelFileNameLabel.Text =
                Path.Combine(Application.StartupPath, EXCEL_FILENAME);

            WriteExcelButton.BackColor = System.Drawing.Color.Yellow;

            //lastDateTimeの初期値を設定
            var now = DateTime.Now;
            //30分以上は30分、30分未満は0分、秒は0秒、ミリ秒は0ミリ秒に調整　2つの方法
            /* 【1】Add～を使用　※分かりにくい
            lastDateTime = now.
            AddMinutes(-now.Minute + ((now.Minute >= 30) ? 30 : 0)).
            AddSeconds(-now.Second).
            AddMilliseconds(-now.Millisecond);
            */
            //【2】new DateTime()で新たに作成
            lastDateTime = new DateTime(
                now.Year, now.Month, now.Day, now.Hour,
                (now.Minute >= 30) ? 30 : 0, 0, 0);
            Debug.WriteLine("lastDateTime: " + lastDateTime);

            timer1_Tick(Type.Missing, EventArgs.Empty); //引数はダミー

            Debug.WriteLine(String.Format("ロケール識別子: [$-{0}]",
                GetSystemDefaultLCID().ToString("x")));
        }

        class Flashair
        {
            MainForm _mainForm;

            public Flashair(MainForm mainForm)
            {
                _mainForm = mainForm;
                _mainForm.FlashairUrlTextBox.Text = "http://";
            }

            public string? Url
            {
                get { return _mainForm.FlashairUrlTextBox.Text; }
                private set { _mainForm.FlashairUrlTextBox.Text = value; }
            }

            public void WriteUrlToInifile()
            {
                if (WritePrivateProfileString(APPNAME, "url",
                    _mainForm.FlashairUrlTextBox.Text, INI_FILENAME))
                {
                    _mainForm.ShowOKMessageBox("保存しました");
                }
                else
                {
                    _mainForm.ShowOKMessageBox("保存できませんでした", CAPTION_ERROR,
                        MessageBoxIcon.Error);
                }
            }

            public void ReadUrlFromInifile()
            {
                //iniファイルから設定を読み込む
                int capacitySize = 256;
                //StringBuilderクラス
                //文字列の追加、置換、挿入を行うと、
                //オブジェクトの内容が変更されるだけで新しいオブジェクトを作成しません
                var sb = new StringBuilder(capacitySize);
                var stringLength = GetPrivateProfileString(
                    APPNAME, "url", "", sb, Convert.ToUInt32(sb.Capacity),
                    INI_FILENAME);
                //FLashAirのURL をラベルに表示
                if (stringLength > 0)
                {
                    _mainForm.FlashairUrlTextBox.Text = sb.ToString();
                }
            }
        }

        class CsvFileList
        {
            MainForm _mainForm;

            public CsvFileList(MainForm mainForm)
            {
                _mainForm = mainForm;
                _mainForm.CsvFileListBox.Items.Clear();
                _mainForm.CsvFileNameLabel.Text = ITEM_NOT_SELECTED;
            }

            public string FileName
            {
                get
                {
                    return _mainForm.CsvFileNameLabel.Text;
                }
                set
                {
                    if (value == null)
                    {
                        return;
                    }
                    _mainForm.CsvFileNameLabel.Text = value;
                }
            }

            public bool Update()
            {
                var list = new List<string>();
                IWebDriver? driver;
                if (_mainForm.ChromeRadioButton.Checked)
                {
                    // Webドライバーのインスタンス化
                    ChromeDriverService? chromeService;
                    ChromeOptions chromeOptions = new();
                    chromeService = ChromeDriverService.CreateDefaultService();
                    //chromeService = ChromeDriverService.CreateDefaultService(Application.StartupPath);
                    //chromeService.SuppressInitialDiagnosticInformation = true; //診断出力抑制
                    chromeService.HideCommandPromptWindow = true; //コマンドプロンプト画面非表示
                    chromeOptions.AddArgument("--headless");
                    //Normal: complete(すべてのリソースをダウンロードするのを待ちます)
                    chromeOptions.PageLoadStrategy = PageLoadStrategy.Normal;

                    try
                    {
                        using (driver = new ChromeDriver(chromeService, chromeOptions))
                        {
                            driver.Navigate().GoToUrl(_mainForm.flashair.Url);
                            ReadOnlyCollection<IWebElement> elms =
                                driver.FindElements(By.XPath(@"//*[@id='thumbnail']/div"));
                            //30分値データのCSVファイル名のリストを取得
                            list = (new List<IWebElement>(elms)).ConvertAll(elm => elm.Text);
                        }
                    }
                    catch (Exception e)
                        when ((e is WebDriverException) || (e is WebDriverArgumentException))
                    {
                        _mainForm.ShowErrorMessageBox("FlashAirのURLが正しいか確認してください");
                        return false;
                    }
                    catch (Exception e)
                    {
                        _mainForm.ShowErrorMessageBox(e);
                        return false;
                    }
                }
                if (_mainForm.EdgeRadioButton.Checked)
                {
                    EdgeDriverService? edgeService;
                    EdgeOptions edgeOptions = new();
                    edgeService = EdgeDriverService.CreateDefaultService();

                    edgeService.HideCommandPromptWindow = true;
                    //上の HideCommandPromptWindow = true でも
                    //コマンドプロンプトが一瞬表示されるため
                    //下の方法(2つ目のAnswer)に変えて無理やり画面外に表示させるようにした
                    //https://stackoverflow.com/questions/35818436/hide-silence-chromedriver-window
                    //edgeOptions.AddArgument("--window-position=-32000,-32000");
                    //Microsoft Edge WebDriver を入れていなかったのが根本原因
                    //Microsoft Edge WebDriver を入れて
                    //HideCommandPromptWindow = true に戻した

                    edgeOptions.AddArgument("--headless");
                    edgeOptions.PageLoadStrategy = PageLoadStrategy.Normal;
                    //edgeOptions.AddArgument("--user-data-dir=C:\\Users\\aida0\\AppData\\Local\\Microsoft\\Edge\\User Data");
                    //edgeOptions.AddArgument("--profile-directory=Default");
                    try
                    {
                        using (driver = new EdgeDriver(edgeService, edgeOptions))
                        {
                            driver.Navigate().GoToUrl(_mainForm.flashair.Url);
                            var elms = driver.FindElements(By.XPath(@"//*[@id='thumbnail']/div"));
                            // ReadOnlyCollection<IWebElement>型のリストから
                            // List<IWebElement>型のリストを生成して
                            // さらに ConvertAll() でList<string>型に変換する
                            list = (new List<IWebElement>(elms)).ConvertAll(elm => elm.Text);
                        }
                    }
                    catch (Exception e)
                        when ((e is WebDriverException) || (e is WebDriverArgumentException))
                    {
                        _mainForm.ShowErrorMessageBox("FlashAirのURLが正しいか確認してください");
                        return false;
                    }
                    catch (Exception e)
                    {
                        _mainForm.ShowErrorMessageBox(e);
                        return false;
                    }
                }
                //リストを空にする
                _mainForm.CsvFileListBox.Items.Clear();
                _mainForm.CsvFileNameLabel.Text = ITEM_NOT_SELECTED;
                if (list.Count == 0)
                {
                    _mainForm.ShowErrorMessageBox("CSVファイルが1つも見つかりませんでした\n指定したURLはFlashAirではない可能性があります");
                    return false;
                }

                //list を条件で絞り込み　2つの方法
                /*
                //【1】for文で処理
                string match;
                foreach (IWebElement elm in elms)
                {
                    match = Regex.Match(elm.Text, @"[0-9]{6}.CSV").ToString();
                    //Debug.WriteLine(match);
                    if (match == String.Empty) //String.Emptyは""と同じ
                    {
                        list.Add(match);
                    }
                }
                */
                //【2】RemoveAll() とラムダ式で処理
                list.RemoveAll(s => !Regex.IsMatch(s, @"[0-9]{6}.CSV"));

                //list を降順ソート　3つの方法 
                //【1】昇順ソートしてから逆順ソート
                //list.Sort();
                //list.Reverse();
                //【2】ラムダ式で降順ソート
                //list.Sort((s1, s2) => s2.CompareTo(s1));
                //【3】LINQで降順ソート
                list.OrderByDescending(s => s);

                //リストに追加
                if (list.Count > 0)
                {
                    _mainForm.CsvFileListBox.Items.AddRange(list.ToArray());
                    _mainForm.CsvFileListBox.SelectedIndex = 0;
                    _mainForm.CsvFileNameLabel.Text = _mainForm.CsvFileListBox.SelectedItem.ToString();
                }
                return true;
            }
        }

        //partial修飾子を付けてProgressForm.csに書かずここに記述
        private partial class ProgressForm : GetFlashairCsv.ProgressForm
        {
            public ProgressForm(System.Drawing.Point point, string text)
            {
                //表示位置の設定
                this.Bounds = new System.Drawing.Rectangle(
                    point.X + 100, point.Y + 150, 240, 100);
                this.Text = text;
                //処理中フォームを表示
                //,Show()でモードレス、.ShowDialog()でモーダル
                this.Show();
                //progressForm のラベルをすぐ表示する　3つの方法
                //【1】メッセージキューに現在あるWindowsメッセージをすべて処理する
                //Application.DoEvents();
                //【2】クライアント領域全体を無効領域に設定し、再描画する
                //progressForm.Update();
                //【3】無効領域(画面更新が必要な領域)を再描画する
                this.Update();
            }
        }

        private bool ExistsForm(Form form)
        {
            return ((form != null) && (form.IsDisposed == false));
        }

        private void ShowErrorMessageBox(Exception e)
        {
            if (ExistsForm(progressForm!))
            {
                progressForm!.Close();
            }
            CustomMessageBox.Show(this, e.ToString(), CAPTION_ERROR,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowErrorMessageBox(string text)
        {
            if (ExistsForm(progressForm!))
            {
                progressForm!.Close();
            }
            CustomMessageBox.Show(this, text, CAPTION_ERROR,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void ShowOKMessageBox(string text,
            string caption = CAPTION_INFORMATION,
            MessageBoxIcon icon = MessageBoxIcon.Information)
        {
            if (ExistsForm(progressForm!))
            {
                progressForm!.Close();
            }
            CustomMessageBox.Show(this, text, caption,
                MessageBoxButtons.OK, icon);
        }

        private DialogResult ShowOKCancelMessageBox(string text,
            string caption = CAPTION_QUESTION,
            MessageBoxIcon icon = MessageBoxIcon.Warning,
            MessageBoxDefaultButton button = MessageBoxDefaultButton.Button1)
        {
            if (ExistsForm(progressForm!))
            {
                progressForm!.Close();
            }
            DialogResult result = CustomMessageBox.Show(mainForm, text, caption,
                MessageBoxButtons.OKCancel, icon, button);
            return result;
        }

        private void WriteInifileButton_Click(object sender, EventArgs e)
        {
            flashair.WriteUrlToInifile();
        }

        private void UpdateCsvFileListButton_Click(object sender, EventArgs e)
        {
            //処理中のフォームを表示
            progressForm = new ProgressForm(this.Location, "リスト更新中...");
            if (csvFileList.Update() == false)
            {
                return;
            }

            //処理中のフォームを閉じる
            progressForm.Close();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CsvFileListBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            // Get the currently selected item in the ListBox.
            csvFileList.FileName = CsvFileListBox.SelectedItem.ToString()!;
        }

        private async Task<bool> DownloadCSVFile()
        {
            //CSVファイル名を取得
            if (CsvFileNameLabel.Text == ITEM_NOT_SELECTED)
            {
                ShowErrorMessageBox("CSVファイルが選択されていません");
                return false;
            }

            //FlashAirのURLにCSVファイル名を連結
            var filepath = FlashairUrlTextBox.Text;
            if (filepath.EndsWith("/") == false)
            {
                filepath += "/";
            }
            filepath += CsvFileNameLabel.Text;

            //CSVファイルをダウンロード
            var client = new HttpClient();
            HttpResponseMessage? response = null;
            try
            {
                response = await client.GetAsync(filepath);
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return false;
            }
            if (response!.StatusCode != System.Net.HttpStatusCode.OK)
            {
                ShowErrorMessageBox("CSVファイルをダウンロードでません");
                return false;
            }
            //保存
            using var stream = await response.Content.ReadAsStreamAsync();
            using var outStream = File.Create(CsvFileNameLabel.Text);
            stream.CopyTo(outStream);
            return true;
        }

        private DialogResult ExistsFile(string fileName)
        {
            //Excelファイルの存在を確認
            var filePath = Path.GetFullPath(fileName);
            Debug.WriteLine("filePath: " + filePath);
            if (File.Exists(filePath))
            {
                return DialogResult.None;
            }
            //Excelファイルが見つからない場合
            //新規作成するかどうか確認
            return ShowOKCancelMessageBox(
                EXCEL_FILENAME + " が見つかりません\n新規作成しますか？");

        }

        private string ConvertDateTimeToString(DateTime date, DateTime time)
        {
            return String.Format("{0} {1}",
                date.ToString(EXCEL_DATE_FORMAT),
                time.ToString(EXCEL_TIME_FORMAT));
        }

        private abstract class ExcelFile
        {
            public static int GetBottomRownum(string range)
            {
                var rownumPart = new Regex(@"[0-9]+$");
                return Int32.Parse(rownumPart.Match(range).ToString());

            }

            //Excelファイルの存在を確認
            //見つかった場合は戻り値Noneを返す
            //見つからない場合、新規作成するならOK、作成しないならCancelを返す
            //
            //MainFormクラスのメンバ変数mainFormをstaticに変えて
            //ShowOKCancelMessageBox()をstaticに変えれば動作はするが
            //staticは使いたくないのでこれは使わない
            /*
            public DialogResult ExistsFile(string fileName)
            {
                //Excelファイルの存在を確認
                string filePath = System.IO.Path.GetFullPath(fileName);
                Debug.WriteLine("filePath: " + filePath);
                if (File.Exists(filePath))
                {
                    return DialogResult.None;
                }
                return ShowOKCancelMessageBox(
                    EXCEL_FILENAME + " が見つかりません\n新規作成しますか？");
            }
            */

            protected DialogResult ExistsFile(MainForm form, string fileName)
            {
                //Excelファイルの存在を確認
                var filePath = Path.GetFullPath(fileName);
                Debug.WriteLine("filePath: " + filePath);
                if (File.Exists(filePath))
                {
                    //Excelファイルが見つかった場合は戻り値Noneを返す
                    return DialogResult.None;
                }
                //Excelファイルが見つからない場合
                //新規作成するならOK、作成しないならCancelを返す
                return form.ShowOKCancelMessageBox(
                    EXCEL_FILENAME + " が見つかりません\n新規作成しますか？");
            }

            public abstract bool Create(string fileName);
            public abstract bool Open(string fileName);
            public abstract StreamReader? OpenCsvFile(string fileName);
            public abstract int Write(StreamReader reader);
            public abstract bool ResizeTable();
            public abstract bool ShowInPane();
            public abstract bool ActivateLastCell();
            public abstract bool Close();
        }

        private class ExcelFileUsingOpenXML : ExcelFile
        {
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
            public ExcelFileUsingOpenXML(MainForm mainForm)
            {
                _document = null;
                _mainForm = mainForm;
            }

            public int GetCount()
            {
                return _count;
            }

            public string ConvertDateTimeToString(DateTime date, DateTime time)
            {
                return String.Format("{0} {1}",
                    date.ToString(EXCEL_DATE_FORMAT),
                    time.ToString(EXCEL_TIME_FORMAT));
            }

            /// <summary>
            /// 列番号からエクセルの列名を得る　(例 5 → "E"）
            /// </summary>
            /// <param name="c"></param>
            /// <returns></returns>
            public string ConvertR1C1ToA1(int r, int c)
            {
                //A～Z列までならこの1行でOK
                //return ((char)(64 + (c + 1))).ToString() + r;

                //string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                var s = "";

                for (; c > 0; c = (c - 1) / 26)
                {
                    var n = (c - 1) % 26;
                    //s = alphabet.Substring(n, 1) + s;
                    s = ((char)(64 + (n + 1))).ToString() + s;
                }
                return s + r;
            }

            public Cell? GetCell(OOXMLS.Worksheet worksheet, string addressName)
            {
                if ((worksheet == null) || (addressName == null))
                {
                    return null;
                }
                //階層構造　Worksheet -> (子)SheetData -> (孫)Row -> (曾孫)Cell
                //worksheet.Elements<cell>()では子要素しか取得できない
                return worksheet.Descendants<Cell>().
                    FirstOrDefault(c => c.CellReference == addressName);
            }

            private string? GetCellValue(WorkbookPart workbookPart, Cell cell)
            {
                if ((workbookPart == null) || (cell == null))
                {
                    return null;
                }
                var text = cell.InnerText;
                if (cell.DataType == null)
                {
                    return null;
                }
                switch (cell.DataType!.Value)
                {
                    case CellValues.SharedString:
                        var index = int.Parse(cell.InnerText);
                        var ssTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                        if (ssTablePart != null)
                        {
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
                int rowMin, int rowMax, int colMin, int colMax)
            {
                if (worksheetPart == null)
                {
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

                var table = new OOXMLS.Table
                {
                    Id = (UInt32)tableNo,
                    Name = EXCEL_TABLENAME,
                    DisplayName = EXCEL_TABLENAME,
                    Reference = reference,
                    TotalsRowShown = false
                };

                var autoFilter = new OOXMLS.AutoFilter() { Reference = reference };

                var tableColumns = new TableColumns()
                {
                    Count = (UInt32)(colMax - colMin + 1)
                };

                //worksheetPartからworkbookPartを取得
                var workbookPart = (WorkbookPart?)worksheetPart.GetParentParts().FirstOrDefault(p => p is WorkbookPart);
                for (int i = 0; i < (colMax - colMin + 1); i++)
                {
                    tableColumns.Append(new TableColumn()
                    {
                        Id = (UInt32)(colMin + i),
                        //Nameは、設定対象のセルに格納されている値と同じ内容を設定する
                        Name = GetCellValue(workbookPart!, 
                            GetCell(worksheetPart.Worksheet, ConvertR1C1ToA1(rowMin, i + 1))!)
                    });
                }

                TableStyleInfo tableStyleInfo = new TableStyleInfo
                {
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
                if (tableParts is null)
                {
                    tableParts = new TableParts();
                    tableParts.Count = (UInt32)0;
                    worksheetPart.Worksheet.Append(tableParts);
                }
                // is not null を != null にするとエラー
                if (tableParts.Count is not null)
                {
                    tableParts.Count += (UInt32)1;
                }
                var tablePart = new TablePart { Id = "rId" + tableNo };

                tableParts.Append(tablePart);

                return table;
            }

            public DialogResult ExistsFile(string fileName)
            {
                return base.ExistsFile(_mainForm, fileName);
            }

            ////Dynamic comments in Excel using Open XML
            //https://social.msdn.microsoft.com/Forums/en-US/c4400c1f-e4b4-43ed-b037-2f531274ea78/dynamic-comments-in-excel-using-open-xml?forum=exceldev
            public bool InsertComments(WorksheetPart worksheetPart,
                List<string> ColumnName, List<string> CellIndex, List<string> NewCommentList)
            {
                if ((ColumnName.Count == 0) || (CellIndex.Count == 0) || (NewCommentList.Count == 0))
                {
                    return false;
                }
                try
                {
                    var commentsVmlXml = string.Empty;
                    // Create all the comment VML Shape XML
                    for (var i = 0; i < ColumnName.Count; i++)
                    {
                        commentsVmlXml += GetCommentVMLShapeXML(ColumnName[i], CellIndex[i]);
                    }
                    var vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>();
                    using (var writer = new XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), Encoding.UTF8))
                    {
                        writer.WriteRaw("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\r\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n <o:shapelayout v:ext=\"edit\">\r\n  <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n" +
                        "</o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\r\n  path=\"m,l,21600r21600,l21600,xe\">\r\n  <v:stroke joinstyle=\"miter\"/>\r\n  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\r\n </v:shapetype>"
                        + commentsVmlXml + "</xml>");
                    }
                    // Create the comment elements
                    for (var j = 0; j < NewCommentList.Count; j++)
                    {
                        var worksheetCommentsPart = worksheetPart.WorksheetCommentsPart ?? worksheetPart.AddNewPart<WorksheetCommentsPart>();
                        // We only want one legacy drawing element per worksheet for comments
                        if (worksheetPart.Worksheet.Descendants<LegacyDrawing>().SingleOrDefault() == null)
                        {
                            string vmlPartId = worksheetPart.GetIdOfPart(vmlDrawingPart);
                            var legacyDrawing = new LegacyDrawing() { Id = vmlPartId };
                            worksheetPart.Worksheet.Append(legacyDrawing);
                        }
                        OOXMLS.Comments comments;
                        bool appendComments = false;
                        if (worksheetPart.WorksheetCommentsPart!.Comments != null)
                        {
                            comments = worksheetPart.WorksheetCommentsPart.Comments;
                        }
                        else
                        {
                            comments = new OOXMLS.Comments();
                            appendComments = true;
                        }
                        // We only want one Author element per Comments element
                        if (worksheetPart.WorksheetCommentsPart.Comments == null)
                        {
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
                            Descendants<OOXMLS.CommentList>().SingleOrDefault() != null))
                        {
                            commentList =
                                worksheetPart.WorksheetCommentsPart.Comments!.
                                Descendants<OOXMLS.CommentList>().Single();
                        }
                        else
                        {
                            commentList = new OOXMLS.CommentList();
                            appendCommentList = true;
                        }
                        var comment = new OOXMLS.Comment()
                        {
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
                        if (appendCommentList)
                        {
                            comments.Append(commentList);
                        }
                        // Only append the Comments if this is the first time adding Comments
                        if (appendComments)
                        {
                            worksheetCommentsPart.Comments = comments;
                        }
                    }
                }
                catch (Exception e)
                {
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
            private string GetCommentVMLShapeXML(string columnName, string rowIndex)
            {
                var commentVmlXml = string.Empty;

                // Parse the row index into an int so we can subtract one
                int commentRowIndex;
                if (int.TryParse(rowIndex, out commentRowIndex) == false)
                {
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
            private string GetAnchorCoordinatesForVMLCommentShape(string columnName, string rowIndex)
            {
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

                if (int.TryParse(rowIndex, out int startingRow) == false)
                {
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
                if (startingRow == 0)
                {
                    coordList[3] = 2;
                    coordList[7] = 16;
                }
                else
                {
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
            public static int? GetColumnIndexFromName(string columnName)
            {
                int? columnIndex = null;

                string[] colLetters = Regex.Split(columnName, "([A-Z]+)");
                colLetters = colLetters.Where(s => !string.IsNullOrEmpty(s)).ToArray();
                var Letters = new List<char>() { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J' };

                if (colLetters.Count() <= 2)
                {
                    int index = 0;
                    foreach (string col in colLetters)
                    {
                        var col1 = colLetters.ElementAt(index).ToCharArray().ToList();
                        var indexValue = Letters.IndexOf(col1.ElementAt(index));

                        if (indexValue != -1)
                        {
                            // The first letter of a two digit column needs some extra calculations
                            if ((index == 0) && (colLetters.Count() == 2))
                            {
                                columnIndex = (columnIndex == null) ?
                                    (indexValue + 1) * 26 :
                                    columnIndex + ((indexValue + 1) * 26);
                            }
                            else
                            {
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

            public override bool Create(string fileName)
            {
                if (String.IsNullOrEmpty(fileName))
                {
                    return false;
                }

                try
                {
                    _document = SpreadsheetDocument.
                        Create(fileName, SpreadsheetDocumentType.Workbook);
                    var workbookPart = _document.AddWorkbookPart();
                    workbookPart.Workbook = new OOXMLS.Workbook();
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var worksheets = _document.WorkbookPart!.Workbook.AppendChild<OOXMLS.Sheets>(new OOXMLS.Sheets());
                    var sheet = new Sheet()
                    {
                        Id = _document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = EXCEL_SHEETNAME
                    };
                    worksheets.Append(sheet);

                    //シートの作成
                    var sheetData = new SheetData();

                    //A列の作成
                    var row = new OOXMLS.Row()
                    {
                        RowIndex = 1U,
                        Spans = new ListValue<OOXML.StringValue>(),
                        Height = 36U,
                        CustomHeight = true
                    };

                    //A1セルの作成
                    Cell cell;
                    cell = new Cell()
                    {
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
                    var font = new OOXMLS.Font()
                    {
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
                    fills.Append(new OOXMLS.Fill()
                    {
                        PatternFill = new OOXMLS.PatternFill
                        {
                            PatternType = PatternValues.None
                        }
                    });
                    fills.Append(new OOXMLS.Fill()
                    {
                        PatternFill = new OOXMLS.PatternFill
                        {
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
                    cellStyleFormats.Append(new OOXMLS.CellFormat()
                    {
                        NumberFormatId = 0,
                        FontId = 0,
                        FillId = 0,
                        BorderId = 0,
                        Alignment = new OOXMLS.Alignment
                        {
                            Vertical = OOXMLS.VerticalAlignmentValues.Center
                        }
                    });
                    stylesheet.Append(cellStyleFormats);

                    //セルフォーマット定義を追加
                    var cellFormats = new OOXMLS.CellFormats();
                    //標準用
                    cellFormats.Append(new OOXMLS.CellFormat()
                    {
                        FormatId = 0,
                        NumberFormatId = 0,
                        FontId = (UInt32Value)0U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        Alignment = new OOXMLS.Alignment()
                        {
                            Vertical = OOXMLS.VerticalAlignmentValues.Center
                        }
                    });
                    //文字折り返し用
                    cellFormats.Append(new OOXMLS.CellFormat()
                    {
                        FormatId = 0,
                        NumberFormatId = 0,
                        FontId = (UInt32Value)0U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        ApplyAlignment = true,
                        Alignment = new OOXMLS.Alignment()
                        {
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
                    SheetView sheetView = new SheetView()
                    {
                        TabSelected = true,
                        WorkbookViewId = (UInt32Value)0U
                    };

                    //ウィンドウ枠の固定の設定
                    //selection1より後にAppendするとエラーになる
                    var pane = new OOXMLS.Pane()
                    {
                        //HorizontalSplit = 1, //列固定
                        VerticalSplit = 1,     //行固定
                        TopLeftCell = "A2",    //下のペインの左上のセル
                        ActivePane = PaneValues.BottomLeft,
                        State = PaneStateValues.Frozen
                    };
                    sheetView.Append(pane);

                    //アクティブセルの設定
                    var selection = new OOXMLS.Selection()
                    {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = "A1",
                        SequenceOfReferences = new ListValue<OOXML.StringValue>()
                        {
                            InnerText = "A1" //R1C1形式は不可
                        }
                    };
                    sheetView.Append(selection);
                    sheetViews.Append(sheetView);

                    worksheet.Append(sheetDimension);
                    worksheet.Append(sheetViews);

                    //ページ設定
                    var sheetFormatProperties = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };
                    var pageMargins = new OOXMLS.PageMargins()
                    {
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
                    lstColumns.Append(new OOXMLS.Column()
                    {
                        Min = 1,
                        Max = 1,
                        Width = 15,
                        CustomWidth = true
                    });
                    lstColumns.Append(new OOXMLS.Column()
                    {
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
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e);
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
            public override bool Open(string fileName)
            {
                if (String.IsNullOrEmpty(fileName))
                {
                    return false;
                }

                try
                {
                    try
                    {
                        _tempFileName = Path.GetTempFileName();
                        Debug.WriteLine("tempFileName: " + _tempFileName);
                        //tempFileName: C:\Users\aida0\AppData\Local\Temp\tmp????.tmp
                    }
                    catch (Exception)
                    {
                        _mainForm.ShowErrorMessageBox("一時ファイルを作成できませんでした");
                        return false;
                    }
                    try
                    {
                        File.Copy(EXCEL_FILENAME, _tempFileName, true);
                    }
                    catch
                    {
                        _mainForm.ShowErrorMessageBox(
                            "Excelファイルを一時ファイルへコピーできませんでした");
                        return false;
                    }
                    _document = SpreadsheetDocument.Open(_tempFileName, true,
                       new OpenSettings { AutoSave = false });
                }
                catch (System.IO.IOException)
                {
                    _mainForm.ShowErrorMessageBox(EXCEL_FILENAME + "にアクセスできません");
                    return false;
                }
                catch (Exception e)
                {
                    _mainForm.ShowErrorMessageBox(e);
                    return false;
                }

                //Excelファイルの中身が正しいか確認
                _workbookPart = _document.WorkbookPart;
                var sheets =
                    _workbookPart!.Workbook.GetFirstChild<OOXMLS.Sheets>();
                _sheet =
                    sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == EXCEL_SHEETNAME);
                if (_sheet == null)
                {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox(EXCEL_SHEETNAME + "が見つかりません");
                    return false;
                }
                Debug.WriteLine("sheet.Name: " + _sheet.Name);
                _worksheetPart = (WorksheetPart)_workbookPart.GetPartById(_sheet.Id!);
                _worksheet = _worksheetPart.Worksheet;
                var a1 = GetCell(_worksheet, "A1");
                if (null == a1)
                {
                    _document.Dispose();
                    File.Delete(_tempFileName);
                    _mainForm.ShowErrorMessageBox("A1セルが見つかりません");
                    return false;
                }
                var cellValue = GetCellValue(_workbookPart, a1);
                Debug.WriteLine("A1: " + cellValue);
                if (cellValue != EXCEL_HEADER_IDENTIFIER)
                {
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
                if (_excelRownum == 1)
                {
                    //ヘッダー行しかない場合
                    _excelDateTime = "";
                    _excelWh = 0.0;
                }
                else
                {
                    var date = GetCell(_worksheet, "A" + _excelRownum);
                    var time = GetCell(_worksheet, "B" + _excelRownum);
                    var wh = GetCell(_worksheet, "C" + _excelRownum);
                    try
                    {
                        _excelDateTime = "";
                        _excelWh = 0.0;
                        //手動でヘッダー行以外を全削除してある場合の対処
                        //テーブルが設定されていると見た目はヘッダー行しかなくても
                        //2行目にRowオブジェクトとCellオブジェクト(InnerTextが"")が存在し
                        //_excelRownumの値が2と認識される
                        if (date!.InnerText == "")
                        {
                            _excelRownum = 1;
                        }
                        else
                        {
                            _excelDateTime = ConvertDateTimeToString(
                                DateTime.FromOADate(double.Parse(date.InnerText)),
                                DateTime.FromOADate(double.Parse(time!.InnerText)));
                            _excelWh = Convert.ToDouble(wh!.InnerText);
                            Debug.WriteLine("ExcelFileUsingOpenXML._excelDateTime: " + _excelDateTime);
                        }
                    }
                    catch (Exception e)
                    {
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

            public override StreamReader? OpenCsvFile(string fileName)
            {
                if ((_document == null) || (String.IsNullOrEmpty(fileName)))
                {
                    return null;
                }
                if (fileName == ITEM_NOT_SELECTED)
                {
                    return null;
                }

                StreamReader reader;
                try
                {
                    var csvFilePath =
                        Path.Combine(Application.StartupPath, fileName);
                    Debug.WriteLine("csvFilePath: " + csvFilePath);
                    reader = new StreamReader(csvFilePath, Encoding.GetEncoding("UTF-8"));
                }
                catch (FileNotFoundException)
                {
                    _document.Dispose();
                    File.Delete(_tempFileName!);
                    _mainForm.ShowErrorMessageBox("CSVファイルが見つかりません");
                    return null;
                }
                catch (Exception e)
                {
                    _document.Dispose();
                    _mainForm.ShowErrorMessageBox(e);
                    return null;
                }

                //CSVファイルのヘッダーを読み込み
                string[] cols = reader.ReadLine()!.Split(',');
                //Excelファイルのヘッダーが仮ヘッダーだったら
                var b1 = GetCell(_worksheet!, "B1");
                if (b1 == null)
                {
                    //CSVファイルのヘッダーをコピー
                    //ヘッダーのセル書式
                    //セル書式が追加済か確認
                    var stylesPart = _workbookPart!.WorkbookStylesPart;
                    var cellFormats = stylesPart!.Stylesheet.CellFormats;
                    object? headerStyleIndex = null;
                    UInt32Value index = 0;
                    foreach (OOXMLS.CellFormat f in cellFormats!)
                    {
                        var alignment = f.Elements<OOXMLS.Alignment>().FirstOrDefault();
                        if ((alignment != null) &&
                            (alignment.WrapText is not null) &&
                            (alignment.WrapText == true))
                        {
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
                    for (int n = 0; n < cols.Length; n++)
                    {
                        cellReference = ConvertR1C1ToA1(1, n + 1);
                        cell = GetCell(_worksheet, cellReference);
                        if (cell == null)
                        {
                            cell = new()
                            {
                                CellReference = cellReference,
                                DataType = CellValues.String,
                                CellValue = new OOXMLS.CellValue(cols[n]),
                            };
                            row.Append(cell);
                        }
                        if (headerStyleIndex != null)
                        {
                            cell.StyleIndex = (UInt32Value)headerStyleIndex;
                        }
                    }
                }
                Debug.WriteLine("ExcelFileUsingOpenXML.OpenCsvFile() 成功");
                return reader;
            }

            public override int Write(StreamReader reader)
            {
                if ((_document == null) || (reader == null))
                {
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
                foreach (OOXMLS.CellFormat f in cellFormats!)
                {
                    if (f.NumberFormatId! == 14)
                    {
                        dateStyleIndex = new UInt32Value(index);
                        break;
                    }
                    index++;
                }
                Debug.WriteLine("dateStyleIndex: " + dateStyleIndex);
                //セル書式を追加
                if (dateStyleIndex == null)
                {
                    stylesPart.Stylesheet.CellFormats!.AppendChild(new OOXMLS.CellFormat()
                    {
                        FormatId = 0,
                        NumberFormatId = 14, //mm-dd-yy
                        FontId = (UInt32Value)0U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        ApplyNumberFormat = OOXML.BooleanValue.FromBoolean(true),
                        ApplyAlignment = true,
                        Alignment = new OOXMLS.Alignment
                        {
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
                if (tempTimeFormat != null)
                {
                    numberingFormat = tempTimeFormat;
                }
                else
                {
                    numberingFormat = new();
                    var tempUserFormat = numberingFormats.
                        Elements<OOXMLS.NumberingFormat>().
                        Where(f => f.NumberFormatId! >= 176).LastOrDefault();
                    if (tempUserFormat != null)
                    {
                        numberingFormat.NumberFormatId =
                            tempUserFormat.NumberFormatId! + 1;
                    }
                    else
                    {
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
                foreach (OOXMLS.CellFormat f in cellFormats)
                {
                    if (f.NumberFormatId! == numberingFormat.NumberFormatId!)
                    {
                        timeStyleIndex = new UInt32Value(index);
                        break;
                    }
                    index++;
                }
                Debug.WriteLine("timeStyleIndex: " + timeStyleIndex);

                //セル書式を追加
                if (timeStyleIndex! == null)
                {
                    //セル書式を作成
                    stylesPart.Stylesheet.CellFormats!.AppendChild(new OOXMLS.CellFormat()
                    {
                        FormatId = (UInt32Value)0U,
                        NumberFormatId = numberingFormat.NumberFormatId,
                        FontId = (UInt32Value)0U,
                        FillId = (UInt32Value)0U,
                        BorderId = (UInt32Value)0U,
                        ApplyNumberFormat = OOXML.BooleanValue.FromBoolean(true),
                        ApplyAlignment = true,
                        Alignment = new OOXMLS.Alignment
                        {
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
                while (reader.Peek() > 0)
                {
                    // 読み込んだ文字列をカンマ区切りで配列に格納
                    cols = reader.ReadLine()!.Split(',');
                    //continue;

                    _csvDateTime = cols[0] + " " + cols[1];
                    //読み込んだデータの日時がExcel最終行より後かどうか判定
                    if (string.Compare(_csvDateTime, _excelDateTime) != 1)
                    {
                        continue;
                    }
                    //Excelのセルオブジェクトの作成とデータ書き込み
                    _excelRownum++;
                    Debug.WriteLine(_excelRownum + "行: " + _csvDateTime);

                    //行追加
                    row = _worksheet!.Descendants<OOXMLS.Row>().FirstOrDefault(r => r.RowIndex! == _excelRownum);
                    if (row == null)
                    {
                        row = new OOXMLS.Row()
                        {
                            RowIndex = Convert.ToUInt32(_excelRownum),
                            //Spans = new ListValue<OOXML.StringValue>()
                        };
                        _sheetData = _worksheet.GetFirstChild<SheetData>();
                        _sheetData!.Append(row);
                    }

                    //1列目 日付
                    date = GetCell(_worksheet, "A" + _excelRownum);
                    if (date == null)
                    {
                        date = new Cell()
                        {
                            CellReference = "A" + _excelRownum,
                            DataType = CellValues.Number,
                            StyleIndex = (UInt32Value)dateStyleIndex,
                        };
                        row.Append(date);
                    }
                    date.CellValue = new OOXMLS.CellValue(DateTime.Parse(cols[0]).ToOADate());

                    //2列目 時刻
                    // 自前でシリアル値に変換
                    var hh = Int32.Parse(cols[1].Substring(0, 2));
                    var mm = Int32.Parse(cols[1].Substring(3));
                    var serialValue = (hh + mm / 60.0) / 24.0;
                    time = GetCell(_worksheet, "B" + _excelRownum);
                    if (time == null)
                    {
                        time = new Cell()
                        {
                            CellReference = "B" + _excelRownum,
                            DataType = CellValues.Number,
                            StyleIndex = (UInt32Value)timeStyleIndex,
                        };
                        row.Append(time);
                    }
                    time.CellValue = new OOXMLS.CellValue(serialValue);

                    //3列目 数値
                    wh = GetCell(_worksheet, "C" + _excelRownum);
                    if (wh == null)
                    {
                        wh = new Cell()
                        {
                            CellReference = "C" + _excelRownum,
                            DataType = CellValues.Number,
                        };
                        row.Append(wh);
                    }
                    wh.CellValue = new OOXMLS.CellValue(cols[2]);

                    //4列目～最終(11列)目 数値
                    for (int n = 3; n < cols.Length; n++)
                    {
                        cell = GetCell(_worksheet, ConvertR1C1ToA1(_excelRownum, n + 1));
                        if (cell == null)
                        {
                            cell = new Cell()
                            {
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

                _mainForm.WriteExcelButton.BackColor = System.Drawing.Color.LightGreen;
                _mainForm.ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                    _excelRownum, _excelDateTime, (int)((_excelWh + 500) / 1000));

                //30分以上は30分、30分未満は0分、秒は0秒、ミリ秒は0ミリ秒に調整　2つの方法
                var dateTime = DateTime.Parse(_excelDateTime!);
                _mainForm.lastDateTime = new DateTime(
                    dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour,
                    (dateTime.Minute >= 30) ? 30 : 0, 0, 0);
                Debug.WriteLine("lastDateTime: " + _mainForm.lastDateTime);

                Debug.WriteLine("ExcelFileUsingOpenXML.Write() 成功");
                return _count;
            }

            public override bool ResizeTable()
            {
                if (_document == null)
                {
                    return false;
                }

                //テーブルの設定
                var tableDefinitionParts = _worksheetPart!.TableDefinitionParts;
                var tableDefinitionPart = tableDefinitionParts.FirstOrDefault(p => p.Table.Name == EXCEL_TABLENAME);
                if (tableDefinitionPart == null)
                {
                    Debug.WriteLine(EXCEL_TABLENAME + "が見つかりません");

                    //テーブルの新規設定
                    if (_worksheetPart.Worksheet.SheetDimension!.Reference == "A1:A1")
                    {
                        _worksheetPart.Worksheet.SheetDimension!.Reference =
                        (OOXML.StringValue)("A1:K" + _excelRownum);
                    }
                    if (AppendTable(_worksheetPart, 1, _excelRownum, 1, 11)
                        == null)
                    {
                        return false;
                    }

                    for (int i = 0; i < (11 - 1 + 1); i++)
                    {
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
                if (_excelRownum != tableLastRownum)
                {
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

            public override bool ShowInPane()
            {
                if (_document == null)
                {
                    return false;
                }

                //最終行を画面内に表示する
                _sheetView = _worksheet!.SheetViews!.GetFirstChild<OOXMLS.SheetView>();
                _pane = _sheetView!.Elements<OOXMLS.Pane>().FirstOrDefault(p => p.ActivePane! == "bottomLeft");
                if (_pane == null)
                {
                    return false;
                }
                var topLeftCell = (_excelRownum < 8) ? _excelRownum : _excelRownum - 8;
                _pane.TopLeftCell = "A" + topLeftCell;
                return true;
            }

            public override bool ActivateLastCell()
            {
                if (_document == null)
                {
                    return false;
                }

                //最終行の1列目のセルをアクティブにする
                _sheetView = _worksheet!.SheetViews!.GetFirstChild<OOXMLS.SheetView>();
                var bottomLeftSelection = _sheetView!.Elements<OOXMLS.Selection>().
                    FirstOrDefault(s => s.Pane! == "bottomLeft");
                if (bottomLeftSelection == null)
                {
                    return false;
                }
                var cell = "A" + _excelRownum;
                bottomLeftSelection.ActiveCell = cell;
                bottomLeftSelection.SequenceOfReferences =
                    new ListValue<OOXML.StringValue>()
                    {
                        InnerText = cell
                    };
                return true;
            }

            public override bool Close()
            {
                if (_document == null)
                {
                    return false;
                }

                //Excelファイル保存
                if (_count > 0)
                {
                    //workbookPart.Workbook.Save(); //これだと不十分
                    _document.Save();
                    _document.Close();
                    try
                    {
                        File.Copy(_tempFileName!, EXCEL_FILENAME, true);
                    }
                    catch (Exception)
                    {
                        _mainForm.ShowErrorMessageBox(
                            "一時ファイルをExcelファイルへコピーできませんでした!!!\n\n" +
                            "<対処方法> 手動で一時ファイル\n" +
                            "(" + _tempFileName + ")を\n" +
                            "Excelファイルに上書きしてから一時ファイルを削除してください");
                        return false;
                    }
                    File.Delete(_tempFileName!);
                }
                else
                {
                    _document.Dispose();
                    File.Delete(_tempFileName!);
                }
                Debug.WriteLine("ExcelFileUsingOpenXML.Close() 成功");
                return true;
            }
        }

        private int WriteExcelUsingOpenXML(MainForm mainForm)
        {
            var excelFile = new ExcelFileUsingOpenXML(mainForm);
            var dialogResult = excelFile.ExistsFile(EXCEL_FILENAME);
            if (dialogResult == DialogResult.Cancel)
            {
                return ERROR_RETURN_VALUE;
            }
            if (dialogResult == DialogResult.OK)
            {
                if (excelFile.Create(EXCEL_FILENAME) == false)
                {
                    return ERROR_RETURN_VALUE;
                }
            }
            if (excelFile.Open(EXCEL_FILENAME) == false)
            {
                return ERROR_RETURN_VALUE;
            }
            var reader = excelFile.OpenCsvFile(CsvFileNameLabel.Text);
            if (reader == null)
            {
                return ERROR_RETURN_VALUE;
            }
            if (excelFile.Write(reader) == ERROR_RETURN_VALUE)
            {
                return ERROR_RETURN_VALUE;
            }
            excelFile.ResizeTable();
            excelFile.ShowInPane();
            excelFile.ActivateLastCell();
            if (excelFile.Close() == false)
            {
                return ERROR_RETURN_VALUE;
            }
            return excelFile.GetCount();
        }

        private async void WriteExcelButton_Click(object sender, EventArgs e)
        {
            if (WriteExcelButton.BackColor == System.Drawing.Color.LightGreen)
            {
                if (ShowOKCancelMessageBox(
                        "Excelファイルは最新です\n実行しますか？",
                        button: MessageBoxDefaultButton.Button2) == DialogResult.Cancel)
                {
                    return;
                }
            }

            //CSVファイルリストが空ならリスト更新
            if (CsvFileListBox.Items.Count == 0)
            {
                progressForm = new ProgressForm(this.Location, "リスト更新中...");
                if (csvFileList.Update() == false)
                {
                    return;
                }
            }

            //CSVファイルをダウンロード
            if (ExistsForm(progressForm!))
            {
                progressForm!.Text = "ダウンロード中...";
            }
            else
            {
                progressForm = new ProgressForm(this.Location, "ダウンロード中...");
            }
            if (await DownloadCSVFile() == false)
            {
                return;
            }

            if (ExistsForm(progressForm!))
            {
                progressForm.Text = "書き込み中...";
            }
            int count;
            //Excelファイルに書き込む　2つの方法
            //【1】NPOIで処理
            //テーブルを操作する方法が見つからない
            //count = WriteExcelUsingNPOI();
            //【2】ClosedXMLで処理
            //最終行を画面内に表示する方法が見つからない
            //count = WriteExcelUsingClosedXML();
            //【3】OpenXMLで処理 ExcelFileUsingOpenXMLクラスで処理
            count = WriteExcelUsingOpenXML(this);

            //ここから【1】【2】【3】共通
            if (count == ERROR_RETURN_VALUE)
            {
                return;
            }
            string message;
            if (count == 0)
            {
                message = "Excelファイルは最新です";
            }
            else
            {
                message = count.ToString() + " 行書き込みました";
            }
            if (ExistsForm(progressForm!))
            {
                progressForm.Close();
            }
            ShowOKMessageBox(message);
        }

        private void OpenExcelUsingProcessStart(string filePath)
        {
            var app = new ProcessStartInfo();
            app.FileName = "excel.exe";
            app.Arguments = filePath;
            app.UseShellExecute = true;
            Process.Start(app);
        }

        private void OpenExcelButton_Click(object sender, EventArgs e)
        {
            //Excelファイルの存在を確認
            string filePath;
            filePath = Path.GetFullPath(EXCEL_FILENAME);
            Debug.WriteLine("filePath: " + filePath);
            if (File.Exists(filePath) == false)
            {
                ShowOKMessageBox(
                    EXCEL_FILENAME + "が見つかりませんでした", CAPTION_ERROR,
                    MessageBoxIcon.Error);
                return;
            }

            //Excelを起動して開く　2つの方法

            //【1】Process.Start() を使用
            OpenExcelUsingProcessStart(filePath);

            //【2】COMオブジェクトを使用
            //テーブル範囲の再設定をWriteExcelButton処理内でしたかったが
            //NPOIにその機能が見つけられなかったためここで処理する
            //OpenExcelUsingCOM(filePath);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            var nowDateTime = DateTime.Now;
            //1秒毎に時計の日時を更新
            Clock.Text = nowDateTime.ToString(CLOCK_FORMAT);

            //30分10秒以内であれば何も処理しない
            if ((nowDateTime - lastDateTime).TotalSeconds <= 1810)
            {
                return;
            }

            if (WriteExcelButton.BackColor != System.Drawing.Color.Orange)
            {
                //WriteExcelButtonのボタンをオレンジ色に
                WriteExcelButton.BackColor = System.Drawing.Color.Orange;

                //アイコン化されていたらサイズを復元　2つの方法
                /*
                //【1】Win32 APIを使ってサイズを復元
                const uint WM_SYSCOMMAND = 0x0112;
                const int SC_RESTORE = 0xF120;

                if (this.WindowState == FormWindowState.Minimized)
                {
                    SendMessage(this.Handle, WM_SYSCOMMAND,
                        new IntPtr(SC_RESTORE), IntPtr.Zero);
                }
                */
                //【2】WindowStateプロパティを使ってサイズを復元
                if (this.WindowState == FormWindowState.Minimized)
                {
                    this.WindowState = FormWindowState.Normal;
                }

                //最前面に
                this.TopMost = true;
                this.TopMost = false;
            }
        }

        //不使用

        private void OpenExcelUsingCOM(string filePath)
        {
            const int HWND_TOPMOST = -1;
            const int HWND_NOTOPMOST = -2;
            const int SWP_NOSIZE = 0x0001;
            const int SWP_NOMOVE = 0x0002;
            const int SWP_SHOWWINDOW = 0x0040;
            ExcelApplication? excel = null;
            Workbooks? workbooks = null;
            Microsoft.Office.Interop.Excel.Workbook? workbook = null;
            Microsoft.Office.Interop.Excel.Sheets? worksheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? worksheet = null;
            Microsoft.Office.Interop.Excel.Window? window = null;
            Microsoft.Office.Interop.Excel.Range? usedRange = null;
            Microsoft.Office.Interop.Excel.Range? usedRows = null;
            Microsoft.Office.Interop.Excel.Range? tableRange = null;
            Microsoft.Office.Interop.Excel.Range? tableRows = null;
            Microsoft.Office.Interop.Excel.Range? range = null;
            Microsoft.Office.Interop.Excel.Range? date = null;
            Microsoft.Office.Interop.Excel.Range? time = null;
            Microsoft.Office.Interop.Excel.Range? wh = null;
            // 既存ファイルオープン
            try
            {
                excel = new ExcelApplication();
                excel.Visible = false; //  Excelを非表示:レスポンス向上
                //workbook = excel.Workbooks.Open(filePath);
                //上のように、下の階層へ2段階以上いっぺんにアクセスすると
                //Excelを終了した段階で中間のリソース(Workbooks)が残ってしまう
                //下のように、1段階ずつリソースにアクセスするようにして
                //必要なくなったら、それぞれをMarshal.ReleaseComObject()する
                workbooks = excel.Workbooks;
                workbook = workbooks.Open(filePath);

                worksheets = workbook.Worksheets;
                Debug.WriteLine("worksheets.Count: " + worksheets.Count);
                int index = 0;
                for (int n = 1; n <= worksheets.Count; n++)
                {
                    worksheet = worksheets[n];
                    if (worksheet.Name == EXCEL_SHEETNAME)
                    {
                        index = n;
                        break;
                    }
                    Marshal.ReleaseComObject(worksheet);

                }
                if (index == 0)
                {
                    ShowOKMessageBox(
                        EXCEL_SHEETNAME + "が見つかりません", "警告",
                        MessageBoxIcon.Warning);
                    worksheet = worksheets[1];
                }
                //1行目にヘッダーがあるかどうか確認
                usedRows = worksheet!.Rows[1];
                date = usedRows.Cells[1];
                if (date.Text.Equals(EXCEL_HEADER_IDENTIFIER) == true)
                {
                    worksheet.Activate();

                    usedRange = worksheet.UsedRange;
                    usedRows = usedRange.Rows;
                    var tables = worksheet!.ListObjects;
                    Debug.WriteLine("tables.Count: " + tables.Count);
                    var usedRowsCount = usedRows.Count;
                    if (tables.Count == 0)
                    {
                        //テーブルを新規作成
                        var table = tables.Add(
                            XlListObjectSourceType.xlSrcRange,
                            usedRange, false, XlYesNoGuess.xlYes, worksheet);
                        table.Name = EXCEL_TABLENAME;
                        table.DisplayName = EXCEL_TABLENAME;
                        table.TableStyle = EXCEL_TABLESTYLENAME;

                        //オートフィルター
                        if (worksheet.AutoFilterMode == false)
                        {
                            worksheet.AutoFilter.ApplyFilter();
                        }

                        //ウィンドウ枠の固定
                        window = excel.ActiveWindow;
                        window.SplitColumn = 0;
                        window.SplitRow = 1;
                        window.FreezePanes = true;

                    }
                    else
                    {
                        //既存のテーブルを探して
                        foreach (ListObject table in tables)
                        {
                            if (table.Name != EXCEL_TABLENAME)
                            {
                                continue;
                            }
                            Debug.WriteLine("table.Name: " + table.Name);

                            //table.Resize()が使えたので
                            //テーブルをセル範囲に変換する table.Unlist();
                            //と、テーブルを新規作成は使わない

                            //テーブル範囲取得
                            tableRange = table.Range;
                            tableRows = tableRange.Rows;
                            int tableRowsCount = tableRows.Count;
                            //テーブル範囲が違っていたら
                            if (usedRowsCount != tableRowsCount)
                            {
                                //テーブル範囲再設定
                                table.Resize(usedRange);
                            }
                        }
                    }

                    //最終行の情報を取得
                    range = worksheet.Rows[usedRowsCount];
                    date = range.Cells[1];
                    time = range.Cells[2];
                    wh = range.Cells[3];
                    var excelLastDateTime = ConvertDateTimeToString(
                        DateTime.FromOADate(date.Value2),
                        DateTime.FromOADate(time.Value2));
                    lastDateTime = DateTime.Parse(excelLastDateTime);

                    //最終行の1列目のセルをアクティブにする
                    date.Activate();

                    //最終行の情報をラベルに表示
                    ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                        usedRowsCount.ToString(), excelLastDateTime,
                        (int)((double.Parse(wh.Text) + 500) / 1000));

                    var nowDateTime = DateTime.Now;
                    if ((nowDateTime - lastDateTime).TotalSeconds > 1810)
                    {
                        if (WriteExcelButton.BackColor != System.Drawing.Color.Orange)
                        {
                            WriteExcelButton.BackColor = System.Drawing.Color.Orange;
                        }
                    }
                    else
                    {
                        if (WriteExcelButton.BackColor != System.Drawing.Color.LightGreen)
                        {
                            WriteExcelButton.BackColor = System.Drawing.Color.LightGreen;
                        }
                    }
                }
                else
                {
                    ShowOKMessageBox(
                        "30分値データが見つかりません", "警告",
                        MessageBoxIcon.Warning);
                }
                excel.Visible = true; // Excel表示
                //Excelを最前面にする
                SetWindowPos((IntPtr)excel.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
                SetWindowPos((IntPtr)excel.Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW | SWP_NOMOVE | SWP_NOSIZE);
            }
            catch (Exception e)
            {
                ReleaseAllComObjects();
                ShowErrorMessageBox(e);
                return;
            }
            ReleaseAllComObjects();
            return;

            //ローカル関数(メソッド内で定義されたメソッド) C#7.0～
            void ReleaseAllComObjects()
            {
                //COM オブジェクトは、アンマネージ リソースのため 
                //自動的に解放されない。解放が漏れるとプロセスが残る。
                //Marshal.ReleaseComObject()を実行することで
                //COM オブジェクトに関連付けられている
                //指定したランタイム呼び出し可能ラッパー(RCW) の
                //参照カウントをデクリメント
                //
                //ReleaseComObject()が漏れるとExcelを閉じてもプロセスが残るが
                //呼び出し元のアプリケーションが終了したタイミングで
                //未廃棄だった.NET オブジェクトの一斉廃棄が起こり、
                //その一環として暗黙参照で使用されていた COM オブジェクトのラッパが解放される
                //結果的に、解放漏れの Excel の COM オブジェクトが解放される
                //逆に、Excelを後から閉じても、閉じたタイミングで解放される
                //<参考>
                //Office オートメーションで割り当てたオブジェクトを解放する - Part1
                //https://learn.microsoft.com/ja-jp/archive/blogs/office_client_development_support_blog/office-5
                if (wh != null) Marshal.ReleaseComObject(wh);
                if (time != null) Marshal.ReleaseComObject(time);
                if (date != null) Marshal.ReleaseComObject(date);
                if (range != null) Marshal.ReleaseComObject(range);
                if (tableRows != null) Marshal.ReleaseComObject(tableRows);
                if (tableRange != null) Marshal.ReleaseComObject(tableRange);
                if (usedRows != null) Marshal.ReleaseComObject(usedRows);
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                if (window != null) Marshal.ReleaseComObject(window);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (worksheets != null) Marshal.ReleaseComObject(worksheets);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                if (excel != null) Marshal.ReleaseComObject(excel);
                // アプリケーションの終了前にガベージ コレクトを強制
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private int WriteExcelUsingClosedXML()
        {
            XLWorkbook? workbook = null;
            IXLWorksheet? worksheet = null;
            IXLTable? table = null;

            //Excelファイルの存在を確認
            var dialogResult = ExistsFile(EXCEL_FILENAME);
            if (dialogResult == DialogResult.Cancel)
            {
                return ERROR_RETURN_VALUE;
            }
            if (dialogResult == DialogResult.OK)
            {
                //ブック新規作成
                workbook = new XLWorkbook();
                //1行目に仮のヘッダー作成
                worksheet = workbook.Worksheets.Add(EXCEL_SHEETNAME);
                worksheet.Cell("A1").Value = EXCEL_HEADER_IDENTIFIER;
            }

            //Excelファイルを開く
            var filePath = Path.GetFullPath(EXCEL_FILENAME);
            Debug.WriteLine("filePath: " + filePath);
            try
            {
                workbook = new XLWorkbook(filePath);
            }
            catch (System.IO.IOException)
            {
                ShowErrorMessageBox(EXCEL_FILENAME + "にアクセスできません");
                return ERROR_RETURN_VALUE;
            }
            catch (System.IO.InvalidDataException)
            {
                ShowErrorMessageBox(EXCEL_FILENAME + "は無効なデータです");
                return ERROR_RETURN_VALUE;
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return ERROR_RETURN_VALUE;
            }

            //目的のシートがあるかどうか確認
            try
            {
                worksheet = workbook.Worksheet(EXCEL_SHEETNAME);
            }
            catch (ArgumentException)
            {
                ShowErrorMessageBox(EXCEL_SHEETNAME + "が見つかりません");
                return ERROR_RETURN_VALUE;
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return ERROR_RETURN_VALUE;
            }

            //1行目にヘッダーがあるかどうか確認
            if (worksheet.Cell(1, 1).Value.Equals(EXCEL_HEADER_IDENTIFIER) == false)
            {
                ShowErrorMessageBox("30分値データが見つかりません");
                return ERROR_RETURN_VALUE;
            }

            //最終行を取得
            var excelRownum = worksheet.RangeUsed().LastRow().RowNumber();
            Debug.WriteLine("excelRownum: " + excelRownum);

            //最終行のデータを取得
            string excelLastDateTime;
            double excelLastWh;
            if (excelRownum == 1)
            {
                //ヘッダー行しかない場合
                excelLastDateTime = "";
                excelLastWh = 0;
            }
            else
            {
                var date = worksheet.Cell(excelRownum, 1);
                var time = worksheet.Cell(excelRownum, 2);
                var wh = worksheet.Cell(excelRownum, 3);
                try
                {
                    excelLastDateTime = ConvertDateTimeToString(
                        date.GetDateTime(), time.GetDateTime());
                    Debug.WriteLine("excelLastDateTime: " + excelLastDateTime);
                    excelLastWh = Convert.ToDouble(wh.Value);
                }
                catch (Exception e)
                {
                    ShowErrorMessageBox(e);
                    return ERROR_RETURN_VALUE;
                }
                ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                    excelRownum, excelLastDateTime,
                    (int)((excelLastWh + 500) / 1000));
            }

            //CSVファイルを開く
            int count = 0;
            int excelLastRownum = 0;
            StreamReader reader;
            try
            {
                string csvFilePath =
                    Path.Combine(Application.StartupPath, CsvFileNameLabel.Text);
                Debug.WriteLine("csvFilePath: " + csvFilePath);
                reader = new StreamReader(csvFilePath, Encoding.GetEncoding("UTF-8"));
            }
            catch (FileNotFoundException)
            {
                ShowErrorMessageBox("CSVファイルが見つかりません");
                return ERROR_RETURN_VALUE;
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return ERROR_RETURN_VALUE;
            }

            string csvDateTime;

            //CSVファイルのヘッダーを読み込み
            string[] cols = reader.ReadLine()!.Split(',');
            //Excelファイルのヘッダーが仮ヘッダーだったら
            if (excelRownum == 1)
            {
                //CSVファイルのヘッダーをコピー
                for (int n = 0; n < cols.Length; n++)
                {
                    worksheet.Cell(1, n + 1).SetValue(cols[n]);
                }
            }

            excelLastRownum = excelRownum;
            while (reader.Peek() > 0)
            {
                // 読み込んだ文字列をカンマ区切りで配列に格納
                cols = reader.ReadLine()!.Split(',');
                csvDateTime = cols[0] + " " + cols[1];
                //読み込んだデータの日時がExcel最終行より後かどうか判定
                if (string.Compare(csvDateTime, excelLastDateTime) != 1)
                {
                    continue;
                }
                //Excelのセルオブジェクトの作成とデータ書き込み
                excelLastRownum++;
                Debug.WriteLine(csvDateTime);

                //1列目 日付
                worksheet.Cell(excelLastRownum, 1).SetValue(DateTime.Parse(cols[0]));
                //～NumberFormat.NumberFormatId = 14ではうまくいかない
                worksheet.Cell(excelLastRownum, 1).Style.NumberFormat.Format = "yyyy/M/d";

                //2列目 時刻
                // 自前でシリアル値に変換
                var hh = Int32.Parse(cols[1].Substring(0, 2));
                var mm = Int32.Parse(cols[1].Substring(3));
                var serialValue = (hh + mm / 60.0) / 24.0;
                worksheet.Cell(excelLastRownum, 2).SetValue(serialValue);
                //～NumberFormat.NumberFormatId = 20ではうまくいかない
                worksheet.Cell(excelLastRownum, 2).Style.NumberFormat.Format =
                    EXCEL_STYLE_TIME_FORMATCODE;

                //3列目～最終(11列)目 数値
                for (int n = 2; n < cols.Length; n++)
                {
                    worksheet.Cell(excelLastRownum, n + 1).
                        SetValue(Convert.ToInt32(cols[n]));
                }
                excelLastDateTime = csvDateTime;
                excelLastWh = worksheet.Cell(excelLastRownum, 3).GetDouble();
                count++;
            }
            reader.Close();
            Debug.WriteLine("CSVファイル読み込み終了");

            //テーブルの設定
            try
            {
                table = worksheet.Tables.Table(EXCEL_TABLENAME);
                Debug.WriteLine("table.Name: " + table.Name);
                var rangeUsed = worksheet.RangeUsed();

                if (table.LastRow().RowNumber() !=
                    rangeUsed.LastRow().RowNumber())
                {
                    //テーブル範囲の再設定
                    table.Resize(rangeUsed);
                }
            }
            catch (ArgumentException)
            {
                ShowErrorMessageBox(
                    EXCEL_TABLENAME + "が見つかりません\r\nテーブルを新規設定します");

                //テーブルの新規設定
                table = worksheet.RangeUsed().AsTable();
                //worksheet.Tables.Add(table) は不要
                table.Name = EXCEL_TABLENAME;
                var themes = XLTableTheme.GetAllThemes();
                //foreach (var theme in themes)
                //{
                //    if (theme.Name == EXCEL_TABLESTYLENAME)
                //    {
                //        table.Theme = theme;
                //    }
                //}
                var theme = themes.FirstOrDefault(t => t.Name == EXCEL_TABLESTYLENAME);
                if (theme != null)
                {
                    table.Theme = theme;
                }
                Debug.WriteLine("table.Theme.Name: " + table.Theme.Name);
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return ERROR_RETURN_VALUE;
            }

            //Excelファイルに書き込み
            if (count > 0)
            {
                //最終行の1列目のセルをアクティブにする
                //worksheet.Cell(excelLastRownum, 1).SetActive();
                //worksheet.Cell(table.LastRow().RowNumber(), 1).Active = true;

                //最終行を画面内に表示する
                //NPOIの ShowInPane() に当たるClosedXMLでの方法を確認中
                // -> 見つからない
                //TopLeftCellAddress はウィンドウの分割だった
                //worksheet.SheetView.TopLeftCellAddress = worksheet.Cell(table.RowCount(), 1).Address;

                //ClosedXMLのオブジェクトとしての保存
                //名前を付けて保存
                try
                {
                    workbook.SaveAs(filePath);
                }
                catch (Exception e)
                {
                    workbook.Dispose();
                    ShowErrorMessageBox(e);
                    return ERROR_RETURN_VALUE;
                }

                //最終行を画面内に表示する
                //ここだけNPOIを使用する
                //try {
                //    var book = new XSSFWorkbook(File.OpenRead(EXCEL_FILENAME));
                //
                //    var sheet = book.GetSheet(EXCEL_SHEETNAME);
                //    Debug.WriteLine("excelLastRownum: " + excelLastRownum);
                //    int toprow = ((excelLastRownum - 1) > 8) ? (excelLastRownum - 1) - 8 : 0;
                //    sheet.ShowInPane(toprow, 0);
                //    sheet.GetRow(excelLastRownum - 1).GetCell(0).SetAsActiveCell();
                //    using (var fs = new FileStream(EXCEL_FILENAME, FileMode.Create, FileAccess.Write))
                //    {
                //        //NPOIのオブジェクトとしての保存
                //        book.Write(fs);
                //    }
                //    book.Close();
                //}
                //               catch (Exception e)
                //{
                //    ShowErrorMessageBox(e);
                //}
                ////保存したファイルをExcelで開くとエラーが発生する
                //　削除されたパーツ: /xl/styles.xml パーツに XML エラーがありました。
                //　(スタイル) 読み込みエラーが発生しました。
                ////NPOIの使用はやめる
                //
            }
            workbook.Dispose();

            lastDateTime = DateTime.Parse(excelLastDateTime);

            WriteExcelButton.BackColor = System.Drawing.Color.LightGreen;
            ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                excelLastRownum, excelLastDateTime,
                (int)((excelLastWh + 500) / 1000));

            return count;
        }

        private int WriteExcelUsingNPOI()
        {
            XSSFWorkbook? workbook = null;
            ISheet worksheet;

            //Excelファイルの存在を確認
            var dialogResult = ExistsFile(EXCEL_FILENAME);
            if (dialogResult == DialogResult.Cancel)
            {
                return ERROR_RETURN_VALUE;
            }
            if (dialogResult == DialogResult.OK)
            {
                //ブック新規作成
                workbook = new XSSFWorkbook();
                workbook.CreateSheet(EXCEL_SHEETNAME);
                //1行目に仮のヘッダー作成
                worksheet = workbook.GetSheet(EXCEL_SHEETNAME);
                var row = worksheet.CreateRow(0);
                var Cell = row.CreateCell(0);
                Cell.SetCellType(NPOI.SS.UserModel.CellType.String);
                Cell.SetCellValue(EXCEL_HEADER_IDENTIFIER);
            }

            //Excelファイルを開く
            try
            {
                //workbookインスタンスを作成した後はこのfsは使わないので破棄
                //usingステートメント使用により自動的にリソースを解放
                using (var fs = File.OpenRead(EXCEL_FILENAME))
                {
                    workbook = new XSSFWorkbook(fs);
                }
            }
            catch (System.IO.IOException)
            {
                ShowErrorMessageBox(EXCEL_FILENAME + "にアクセスできません");
                return ERROR_RETURN_VALUE;
            }
            catch (System.IO.InvalidDataException)
            {
                ShowErrorMessageBox(EXCEL_FILENAME + "は無効なデータです");
                return ERROR_RETURN_VALUE;
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return ERROR_RETURN_VALUE;
            }

            //目的のシートがあるかどうか確認
            worksheet = workbook.GetSheet(EXCEL_SHEETNAME);
            if (worksheet == null)
            {
                ShowErrorMessageBox(EXCEL_SHEETNAME + " が見つかりません");
                return ERROR_RETURN_VALUE;
            }

            //1行目にヘッダーがあるかどうか確認
            if (worksheet.GetRow(0).GetCell(0).StringCellValue.
                Equals(EXCEL_HEADER_IDENTIFIER) == false)
            {
                ShowErrorMessageBox("30分値データが見つかりません");
                return ERROR_RETURN_VALUE;
            }

            //最終行のデータを取得
            var excelRownum = worksheet.LastRowNum;
            string excelLastDateTime;
            double excelLastWh;
            if (excelRownum == 0)
            {
                Debug.WriteLine(excelRownum);
                //ヘッダー行しかない場合
                excelLastDateTime = "";
                excelLastWh = 0;
            }
            else
            {
                var row = worksheet.GetRow(excelRownum);
                var date = row.GetCell(0);
                var time = row.GetCell(1);
                var wh = row.GetCell(2);
                try
                {
                    excelLastDateTime = ConvertDateTimeToString(
                        date.DateCellValue, time.DateCellValue);
                    excelLastWh = wh.NumericCellValue;
                }
                catch (Exception e)
                {
                    ShowErrorMessageBox(e);
                    return ERROR_RETURN_VALUE;
                }
                ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                    (excelRownum + 1).ToString(), excelLastDateTime,
                    (int)((excelLastWh + 500) / 1000));
            }
            int excelLastRownum = excelRownum;

            //CSVファイルを開く
            int count = 0;
            StreamReader reader;
            try
            {
                var csvFilePath =
                    Path.Combine(Application.StartupPath, CsvFileNameLabel.Text);
                Debug.WriteLine(csvFilePath);
                reader = new StreamReader(csvFilePath, Encoding.GetEncoding("UTF-8"));
            }
            catch (FileNotFoundException)
            {
                ShowErrorMessageBox("CSVファイルが見つかりません");
                return ERROR_RETURN_VALUE;
            }
            catch (Exception e)
            {
                ShowErrorMessageBox(e);
                return ERROR_RETURN_VALUE;
            }

            var dateStyle = workbook.CreateCellStyle();
            dateStyle.DataFormat =
                workbook.CreateDataFormat().GetFormat(EXCEL_STYLE_DATE_FORMATCODE);
            var timeStyle = workbook.CreateCellStyle();
            timeStyle.DataFormat =
                workbook.CreateDataFormat().GetFormat(EXCEL_STYLE_TIME_FORMATCODE);
            string csvDateTime;
            IRow excelRow;
            NPOI.SS.UserModel.ICell excelCell;

            //1行目を空読み
            string[] cols = reader.ReadLine()!.Split(',');
            //Excelファイルのヘッダーが仮ヘッダーだったら
            if (excelRownum == 0)
            {
                //CSVファイルのヘッダーをコピー
                excelRow = worksheet.CreateRow(0);
                for (int n = 0; n < cols.Length; n++)
                {
                    excelCell = excelRow.CreateCell(n);
                    excelCell.SetCellType(NPOI.SS.UserModel.CellType.String);
                    excelCell.SetCellValue(cols[n]);
                }
            }

            while (reader.Peek() > 0)
            {
                // 読み込んだ文字列をカンマ区切りで配列に格納
                cols = reader.ReadLine()!.Split(',');
                csvDateTime = cols[0] + " " + cols[1];
                //読み込んだデータの日時がExcel最終行より後かどうか判定
                if (string.Compare(csvDateTime, excelLastDateTime) != 1)
                {
                    continue;
                }
                //Excelのセルオブジェクトの作成とデータ書き込み
                excelLastRownum++;
                Debug.WriteLine(csvDateTime);
                excelRow = worksheet.CreateRow(excelLastRownum);

                //1列目 日付
                excelCell = excelRow.CreateCell(0);
                excelCell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                excelCell.CellStyle = dateStyle;
                //Parse()だと月日年になる時があったので ParseExact()に変更
                excelCell.SetCellValue(DateTime.ParseExact(
                    cols[0], EXCEL_STYLE_DATE_FORMATCODE, null));

                //2列目 時刻
                excelCell = excelRow.CreateCell(1);
                excelCell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                excelCell.CellStyle = timeStyle;
                // 自前でシリアル値に変換
                var hh = Int32.Parse(cols[1].Substring(0, 2));
                var mm = Int32.Parse(cols[1].Substring(3));
                var serialValue = (hh + mm / 60.0) / 24.0;
                excelCell.SetCellValue(serialValue);

                //3列目～最終(11列)目 数値
                for (int n = 2; n < cols.Length; n++)
                {
                    excelCell = excelRow.CreateCell(n);
                    excelCell.SetCellType(NPOI.SS.UserModel.CellType.Numeric);
                    excelCell.SetCellValue(Int32.Parse(cols[n]));
                }
                excelLastDateTime = csvDateTime;
                excelLastWh = excelRow.GetCell(2).NumericCellValue;
                count++;
            }
            reader.Close();
            Debug.WriteLine("CSVファイル読み込み終了");

            //テーブル範囲の設定
            //処理方法見つからず
            //ここだけClosedXMLで処理して試したがうまくいかない

            //Excelファイルに書き込み
            if (count > 0)
            {
                //最終行の1列目のセルをアクティブにする
                //いつの間にか効かなくなっていた
                //いつの間にかXSSFではなくHSSFを参照?
                //参考
                //poi。HSSFCell#setAsActiveCellの問題。
                //https://tak-tech.hatenablog.com/entry/20121005/1349403896
                //worksheet.GetRow(excelLastRownum).GetCell(0).SetAsActiveCell();
                //これも効かない
                //worksheet.SetActiveCell(excelLastRownum, 0);
                //これは未実装でビルドエラー
                //worksheet.SetActiveCellRange(excelLastRownum, excelLastRownum, 0, 0);
                //XSSFでブックを操作したら効くようになった
                worksheet.GetRow(excelLastRownum).GetCell(0).SetAsActiveCell();

                //最終行を画面内に表示する
                var toprow = (excelLastRownum > 8) ? excelLastRownum - 8 : 0;
                worksheet.ShowInPane(toprow, 0);

                try
                {
                    //Writeメソッドを実行した後はこのfsは使わないので破棄
                    //usingステートメント使用により自動的にリソースを解放
                    using (var fs = new FileStream(EXCEL_FILENAME, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fs);
                    }
                }
                catch (Exception e)
                {
                    workbook.Close();
                    ShowErrorMessageBox(e);
                    return ERROR_RETURN_VALUE;
                }
            }
            workbook.Close();
            lastDateTime = DateTime.Parse(excelLastDateTime);
            WriteExcelButton.BackColor = System.Drawing.Color.LightGreen;
            ExcelLastDataLabel.Text = string.Format("{0}行 {1} {2}kWh",
                (excelLastRownum + 1).ToString(), excelLastDateTime,
                (int)((excelLastWh + 500) / 1000));
            return count;
        }
    }
}