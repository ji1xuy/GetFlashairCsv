# GetFlashairCsv
■概要  
　電力モニター( https://github.com/ji1xuy/WattMonitor )に内蔵のメモリーカードのCSVデータを、Wi-Fi越しにパソコンで取得してExcelファイルに蓄積するツール  
■機能  
・電力モニターのIPアドレスを設定ファイルから読み出し・保存  
・電力モニター内のCSVファイル名をリスト化  
・電力モニターからCSVファイルをダウンロード  
・ダウンロードしたCSVファイルを読み込み、増分をExcelファイルに追加(※)  
・Excelファイルがない場合は、新規作成(※)　　　　※Excel不要  
・ファイルをExcelで開く  
  
■Visual Studioでの準備  
NPOI, COM, ClosedXML, OpenXMLすべて必要ではありません。一つの機能をそれぞれの方法で記述し残してあります。MainForm.csの2053行～2799行のコードを削除すれば、NPOI, COM, ClosedXMLはインストール不要です。  
  
・Nuget パッケージのインストール  
　DotNetCore.NPOI  
　Selenium.WebDriver  
　Selenium.WebDriver.ChromeDriver  
　ClosedXML  
　DocumentFormat.OpenXml  
・COM参照の追加  
　Microsoft Excel 16.0 ObjectLibrary  
・Microsoft Edge WebDriverのダウンロード  
　https://developer.microsoft.com/ja-jp/microsoft-edge/tools/webdriver/  
　  
■スクリーンショット  
  ![screenshot](https://user-images.githubusercontent.com/114241917/209545848-695872a3-73ca-4d4a-8c29-51de413dd6b2.jpg)
