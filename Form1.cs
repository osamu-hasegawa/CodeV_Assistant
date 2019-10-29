using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Collections;
using System.Diagnostics; 

namespace CodeV_Assistant
{
    public partial class Form1 : Form
    {
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern int GetWindowText(int hWnd,
		    StringBuilder lpString, int nMaxCount);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern int GetWindowTextLength(int hWnd);

		[DllImport("user32.dll", SetLastError = true)]
		static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, int wParam, StringBuilder lParam);

		[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
		private static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

		[DllImport("user32.dll", SetLastError = true)]
		static extern int GetWindow(int hWnd, int uCmd);

        public const int GW_HWNDNEXT = 2;
        public const int GW_CHILD = 5;
		public const uint WM_SETTEXT = 0x000C;

		public const uint WM_KEYDOWN = 0x0100;
		public const uint WM_KEYUP = 0x0101;
		public const uint VK_RETURN = 0x000D;

		public int destWnd = 0;
		public FileInfo[] fi;

        public Form1()
        {
            InitializeComponent();
			button1.Enabled = true;
			button2.Enabled = false;
			label1.Text = "";
			label2.Text = "";
			label3.Text = "";
        }

		public static void SendText(int hWnd, string text)
		{

		    StringBuilder sb = new StringBuilder(text);
		    SendMessage((IntPtr)hWnd, WM_SETTEXT, (int)0, sb);
		}

		public bool SearchWindow()
		{
            int topWnd = 0;
			//プロセスよりCODE Vのメインハンドル取得
            System.Diagnostics.Process[] process;
            //"CODE V"のタスクマネージャー上の名称[.exe]除く。　からウィンドウハンドルを取得
            process = System.Diagnostics.Process.GetProcessesByName("cvgui");
            foreach (System.Diagnostics.Process ps in process)
            {
                topWnd = (int)ps.MainWindowHandle;//CODE Vのメインハンドル
            }

			if(topWnd == 0)//CODE Vが起動していない場合等
			{
				return false;
			}

			//メインハンドルから子ウインドウ、兄弟ウインドウを手繰り寄せる
            int childWnd = FindWindowEx(topWnd, 0, "MDIClient", "");//MDIウインドウ

			//MDIの中のフォーカスがあるウィンドウが取得される。そのウィンドウのうち、コマンドウインドウを見つける
            int hChild = GetWindow(childWnd, GW_CHILD);
			do
			{
				//ウィンドウのタイトルの長さを取得する
				int textLen = GetWindowTextLength(hChild);
				if (textLen > 0)
				{
		            //ウィンドウのタイトルを取得する
		            StringBuilder tsb = new StringBuilder(textLen + 1);
		            GetWindowText(hChild, tsb, tsb.Capacity);

                    if (tsb.ToString() == "コマンドウインドウ")
                    {
                        break;
                    }
                }

                hChild = GetWindow(hChild, GW_HWNDNEXT);

		    }
		    while(hChild != 0);

			int hChildChild = GetWindow(hChild, GW_CHILD);//AfxMDIFrame120u
			int hGrandchild = GetWindow(hChildChild, GW_HWNDNEXT);//AfxWnd120u
			int hComboBox = GetWindow(hGrandchild, GW_CHILD);//ComboBox
			int hEditBox = GetWindow(hComboBox, GW_CHILD);//EditBox
			destWnd = hEditBox;
			
			return true;
		}


		public void ControlCodeV()
		{
			//lensファイル一覧を取得
			string dirpath = @"C:\CVUSER";
			DirectoryInfo di = new DirectoryInfo(dirpath);
			fi = di.GetFiles(@"lens*.len");

			string baseFilename = "macro_base.txt";
			string workFilename = "macro_work.txt";
			string basePath = dirpath + "\\" + baseFilename;
			string workPath = dirpath + "\\" + workFilename;
			string oldtext = "XXXXXXXX";
			string newtext = "";

			Stopwatch sw = new Stopwatch(); 
			sw.Start(); 

            for (int j = 0; j < fi.Length; j++)
            {
				// キャンセル通知があればスレッドをキャンセルする
				if(backgroundWorker1.CancellationPending)
				{
					return;
				}

				//進捗率をUIスレッドに送信
                int prog = (int)(((double)(j + 1) / (double)fi.Length) * 100f);//[%]となる
				string status = string.Format("{0}ファイル／全{1}ファイル", (j + 1), fi.Length);

				TimeSpan ts = sw.Elapsed;
				string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}", ts.Hours, ts.Minutes, ts.Seconds);

                string msg = fi[j].Name + " " + status + " " + elapsedTime;
                backgroundWorker1.ReportProgress(prog, msg);

				//ファイルがあれば削除する
				if(System.IO.File.Exists(@workPath))
				{
					System.IO.File.Delete(@workPath);
				}

                newtext = fi[j].Name;
                
                StringBuilder strread = new StringBuilder();
                string[] strarray = File.ReadAllLines(basePath, Encoding.GetEncoding("SHIFT_JIS"));
				int count = 0;
				string noExt = "";
				string replace = "";
                for (int i = 0; i < strarray.GetLength(0); i++)
                {
                    if(strarray[i].Contains(oldtext))
                    {
						//ファイル名から拡張子を除く
						noExt = Path.GetFileNameWithoutExtension(newtext); 
						if(count == 0)//lensファイル名そのままに置き換える
						{
	                        strread.AppendLine(strarray[i].Replace(oldtext, noExt));
	                    }
	                    else
	                    {
							replace = "";
							if(noExt.IndexOf(".") > 0)
							{
								replace = noExt.Replace(".", "_");//ファイル名に"."が含まれていた場合は"_"に置き換える
							}
							else
							{
								replace = noExt;
	                        }
                        	strread.AppendLine(strarray[i].Replace(oldtext, replace));
						}
						count++;
                    }
                    else
                    {
                        strread.AppendLine(strarray[i]);
                    }
                }
				File.WriteAllText(workPath, strread.ToString(), Encoding.GetEncoding("SHIFT_JIS"));


				//ファイルがあれば削除する
				if(System.IO.File.Exists(@dirpath + "\\" + replace + "_macro.txt"))
				{
					System.IO.File.Delete(@dirpath + "\\" + replace + "_macro.txt");
				}

                Thread.Sleep(100);//念の為

				File.Copy(@workPath, @dirpath + "\\" + replace + "_macro.txt");

                //マクロファイルを読みこむ
                FileStream fs = new FileStream(workPath, FileMode.Open, FileAccess.Read);
                StreamReader sr = new StreamReader(fs, Encoding.GetEncoding("SHIFT_JIS"));
                try
                {
                    while(sr.EndOfStream == false)
                    {
                        string line = sr.ReadLine();
                        if (line != "")
                        {
                            SendText(destWnd, line);//CODE Vに1行送信
                            Thread.Sleep(20);
                            PostMessage((IntPtr)destWnd, WM_KEYDOWN, (int)VK_RETURN, 0);//Enterキーを押下
                            Thread.Sleep(20);
                            PostMessage((IntPtr)destWnd, WM_KEYUP, (int)VK_RETURN, 0);//Enterキーを戻す
                            Thread.Sleep(20);
                        }
                    }
                }
                catch (Exception exc)
                {
                }
                finally
                {
                    sr.Close();
                    fs.Close();
                }

                Thread.Sleep(1000);//次のlensファイルまでのインターバル
            }
			
			sw.Stop(); 
        }

        private void button1_Click(object sender, EventArgs e)
        {
			button1.Enabled = false;
			button2.Enabled = true;

            if(!SearchWindow())
            {
				button1.Enabled = true;
				button2.Enabled = false;

				MessageBox.Show("CODE Vが起動していない可能性があります。確認して下さい。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

				return;
			}

            backgroundWorker1.RunWorkerAsync(0);
        }

        private void button2_Click(object sender, EventArgs e)
        {
			button1.Enabled = true;
			button2.Enabled = false;

            backgroundWorker1.CancelAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
			ControlCodeV();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;

            string line = (string)e.UserState;
            string[] stringValues = line.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            label1.Text = stringValues[0];
            label2.Text = stringValues[1];
            label3.Text = "経過時間：" + stringValues[2] + "[s]";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                label1.Text = "キャンセルされました";
                progressBar1.Value = 0;
                return;
            }

            label1.Text = "終了しました";
			button1.Enabled = true;
			button2.Enabled = false;
        }
    }
}
