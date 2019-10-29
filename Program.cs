using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;

namespace CodeV_Assistant
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //ミューテックス作成
            Mutex app_mutex = new Mutex(false, "CodeV_Assistant");

            //ミューテックスの所有権を要求する
            if (app_mutex.WaitOne(0, false) == false)
            {
                MessageBox.Show("すでに、このアプリは起動しています。");
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
