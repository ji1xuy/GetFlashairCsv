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
・Nuget パッケージのインストール  
　Selenium.WebDriver  
　WebDriverManager  
　DocumentFormat.OpenXml  

■スクリーンショット  
　![screenshot](https://github.com/ji1xuy/GetFlashairCsv/assets/114241917/28f9f840-a5f0-4b0b-b3c3-b0e03731aba7)

　
