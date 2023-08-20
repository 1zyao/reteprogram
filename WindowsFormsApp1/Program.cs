using System.Windows.Forms;
using System;

namespace WinFormsApp1
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            WelcomeForm welcomeForm = new WelcomeForm();
            Application.Run(welcomeForm); // 显示 WelcomeForm 窗体

            if (welcomeForm.DialogResult == DialogResult.OK)
            {
                MainForm mainForm = new MainForm();
                Application.Run(mainForm); // 显示 MainForm 窗体
            }
        }
    }
}