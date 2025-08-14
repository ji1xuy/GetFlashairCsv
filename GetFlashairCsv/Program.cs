namespace GetFlashairCsv
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //ミューテックス作成
            Mutex mutex = new Mutex(false, "GET_FLASHAIR_CSV");

            //ミューテックスの所有権を要求する
            if (mutex.WaitOne(0, false) == false)
            {
                if (MessageBox.Show("既に起動しています!\n続行しますか？", "確認",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning
                    ) == DialogResult.Cancel)
                {
                    return;
                }
            }
            
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());

        }
    }
}