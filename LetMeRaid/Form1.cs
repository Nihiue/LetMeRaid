using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;

namespace LetMeRaid
{
    public partial class MainWindow : Form

    {
        private System.Timers.Timer tickTimer;
        private bool enableAutoRestart = true;
        delegate void deleAppendLog(string text);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll ")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);


        public static Bitmap getScreen()
        {
            Bitmap baseImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(baseImage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), Screen.AllScreens[0].Bounds.Size);
            g.Dispose();
            return baseImage;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Timers.Timer timer = new System.Timers.Timer(5000);
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.onTick);
            timer.AutoReset = true;
            timer.Enabled = false;
            this.tickTimer = timer;
        }

        public void onTick(object source, System.Timers.ElapsedEventArgs e) {
            DateTime dt = DateTime.Now;
            bool inTimeRange = dt.TimeOfDay.TotalMinutes >= (double)(this.numericUpDown1.Value * 60 + this.numericUpDown2.Value);

            if (!inTimeRange) {
                return;
            }

            bool[] psStatus = this.checkProcess();

            if (!psStatus[0]) {
                this.BeginInvoke(new deleAppendLog(appendLog), "未检测到战网客户端");
                return;
            }
            if (psStatus[1])
            {
                int status = this.getWowStatus();

                if (status == -1 && this.enableAutoRestart) {                    
                    this.BeginInvoke(new deleAppendLog(appendLog), "检测到掉线，关闭魔兽世界");
                    this.killWow();
                } else if (status == 1) {
                    this.BeginInvoke(new deleAppendLog(appendLog), "检测到人物选择界面，选择人物");
                    Thread thread = new Thread(new ThreadStart(enterGame));
                    thread.IsBackground = true;
                    thread.Start();
                }
            }
            else {
                this.BeginInvoke(new deleAppendLog(appendLog), "启动魔兽世界");
                Thread thread = new Thread(new ThreadStart(launchWow));
                thread.IsBackground = true;
                thread.Start();
            }
        }

        /*
            -1: 被踢
            0: 正常
            1: 角色选择
        */

        private int calColorErr(Color c1, Color c2)
        {
            return (c1.R - c2.R) * (c1.R - c2.R) + (c1.G - c2.G) * (c1.G - c2.G) + (c1.B - c2.B) * (c1.B - c2.B);
        }

        private long getDeltaE(LockBitmap lockbmp, int[,] pts) {

            Color st = Color.FromArgb(123, 9, 6);
            long e = 0;

            int width = lockbmp.Width;
            int height = lockbmp.Height;

            for (int i = 0; i < 2; i++) {
                int ptx = pts[i,0] * width / 2560;
                int pty = pts[i,1] * height / 1440;
                if (e > 1000000) {
                    break;
                }
                for (int j = -2; j < 3; j++) {
                    e += calColorErr(st, lockbmp.GetPixel(ptx + j, pty + j));
                }
            }

            return e;
        }
        private int getWowStatus() {
            int[,] loginPts = {
                { 70, 1035 }, 
                { 235, 1034 },
                { 77, 1106 },
                { 235, 1106 },
                { 75, 1176 },
                { 237, 1176 },
                { 1211, 1013 },
                { 1381, 1013 },
                { 2332, 990 },
                { 2327, 1057 },
                { 2327, 1125 },
                { 2337, 1345 }
            };

            int[,] choosePts = {
                { 118, 1355 },
                { 262, 1355 },
                { 260, 1358 },
                { 267, 1407 },
                { 1155, 1323 },
                { 1396, 1321 },
                { 2219, 97 },
                { 2188, 1155 },
                { 2070, 1357 },
                { 2238, 1356 },
                { 2362, 1358 },
                { 2462, 1358 }
            };

            if (this.activateWindow("魔兽世界"))
            {
               Bitmap bmp = getScreen();
               LockBitmap lockbmp = new LockBitmap(bmp);
               lockbmp.LockBits();

               if (getDeltaE(lockbmp, loginPts) < 10000)
               {
                   return -1;
               }
               if (getDeltaE(lockbmp, choosePts) < 10000)
               {
                   return 1;
               }
            }
            else
            {
                Console.WriteLine("WOW NOT FOUND");          
            }
            return 0;
        }

        private void killWow() {
            foreach (Process p in Process.GetProcesses())
            {
                if (p.ProcessName.ToLower() == "wow")
                {
                    try
                    {
                        p.Kill();
                        p.WaitForExit();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message.ToString());
                    }
                }

            }
        }

        private void enterGame() {
            int cx = Screen.PrimaryScreen.Bounds.Width / 2;
            int cy = Screen.PrimaryScreen.Bounds.Height * 1320 / 1440;
            SetCursorPos(cx, cy);
            mouse_event(0x0002, 0, 0, 0, 0); 
            Thread.Sleep(55);
            mouse_event(0x0004, 0, 0, 0, 0); 
            Thread.CurrentThread.Abort();
        }

        private bool activateWindow(string title) {
            IntPtr findPtr = FindWindow(null, title);

            if (findPtr.ToInt32() != 0) {
                ShowWindow(findPtr, 9);
                SetForegroundWindow(findPtr);
                return true;
            }
            else {
                return false;
            }
        }

        private void launchWow() {
            if (this.activateWindow("暴雪战网"))
            {
                Thread.Sleep(300);
                SendKeys.SendWait("{ENTER}");
            }
            else {
                Console.WriteLine("BN NOT FOUND");
            }
            Thread.CurrentThread.Abort();
        }

        private bool[] checkProcess() {
            bool[] ret = { false, false };
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps) {
                if (p.ProcessName.ToLower() == "battle.net") {
                    ret[0] = true;
                }
                if (p.ProcessName.ToLower() == "wow")
                {
                    ret[1] = true;
                }
            }
                return ret;
        }



        private void appendLog(string log) {
            List<string> tmp = this.textBox1.Lines.ToList();
            DateTime dt = DateTime.Now;
            tmp.Add(string.Format("{0:T}", dt) + " " + log);
            this.textBox1.Lines = tmp.ToArray();
        }

        private void toggleUI(bool running) {
            this.button1.Enabled = !running;
            this.button2.Enabled = running;
            this.radioButton1.Enabled = !running;
            this.radioButton2.Enabled = !running;
            this.numericUpDown1.Enabled = !running;
            this.numericUpDown2.Enabled = !running;
        }
        private void startService() {
            this.toggleUI(true);
            this.enableAutoRestart = this.radioButton1.Checked;
            this.tickTimer.Enabled = true;
            appendLog("启用服务");
        }
        private void stopService() {
            this.toggleUI(false);
            this.tickTimer.Enabled = false;
            appendLog("停止服务");
        }
        private void button1_Click(object sender, EventArgs e)
        {
            bool[] psStatus = this.checkProcess();
            if (!psStatus[0])
            {
                MessageBox.Show("请先运行战网客户端！");
                return;
            }
            this.startService();          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.stopService();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("使用说明\n\n1. 保持战网开启并激活魔兽世界子页面\n2. 魔兽世界配置为窗口（最大化）模式\n\nBY:小脑斧\nGithub:Nihiue");
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
    }
}
