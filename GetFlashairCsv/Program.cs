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
            //�~���[�e�b�N�X�쐬
            Mutex mutex = new Mutex(false, "GET_FLASHAIR_CSV");

            //�~���[�e�b�N�X�̏��L����v������
            if (mutex.WaitOne(0, false) == false)
            {
                if (MessageBox.Show("���ɋN�����Ă��܂�!\n���s���܂����H", "�m�F",
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