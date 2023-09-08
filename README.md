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
  
■Visual Studioでの準備  ・Nuget パッケージのインストール　 
　DotNetCore.NPOI  
　Selenium.WebDriver  
　WebDriverManager  
　ClosedXML  
　DocumentFormat.OpenXml  
・COM参照の追加  
　Microsoft Excel 16.0 ObjectLibrary  

NPOI, COM, ClosedXML, OpenXMLすべて必要ではありません。一つの機能をそれぞれの方法で記述し残してあります。MainForm.csの2154行から最終行までのコードを削除すれば、NPOI, COM, ClosedXMLはインストール不要です。それ以外にも、本来は必要のない、興味本位で作ったコード、後で使い回しするかもしれないと思ったコードが残してあります。それらは冗長なコードなので削っても支障はありません。

■スクリーンショット  
　![screenshot](https://github.com/ji1xuy/GetFlashairCsv/assets/114241917/28f9f840-a5f0-4b0b-b3c3-b0e03731aba7)

　
