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
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetClientRect(IntPtr hWnd, ref RECT lpRect);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left; 
            public int Top; 
            public int Right; 
            public int Bottom;
        }

        [DllImport("user32")]
        public static extern int GetSystemMetrics(int nIndex);
        [DllImport("user32", EntryPoint = "HideCaret")]
        private static extern bool HideCaret(IntPtr hWnd);
        private static Point getClientStart(ref RECT cRect, ref RECT wRect) {
            Point start = new Point(wRect.Left, wRect.Top);
            if (wRect.Bottom - wRect.Top != cRect.Bottom)
            {
                // 非全屏                
                int titleHeight = GetSystemMetrics(4);
                int borderWidth = (wRect.Right - wRect.Left - cRect.Right) / 2;
                start.X = start.X + borderWidth;
                start.Y = start.Y + titleHeight + borderWidth;
            }
            return start;
        }
        public static Bitmap getScreenshot(ref RECT cRect, ref RECT wRect)
        {
           
            Size capSize = new Size(cRect.Right - cRect.Left, cRect.Bottom - cRect.Top);
            Point capStart = getClientStart(ref cRect, ref wRect);
            Bitmap baseImage = new Bitmap(capSize.Width, capSize.Height);
            Graphics g = Graphics.FromImage(baseImage);
            g.CopyFromScreen(capStart, new Point(0, 0), capSize);
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

            if (!inTimeRange && false) {
                return;
            }

            bool[] psStatus = this.checkProcess();

            if (!psStatus[0]) {
                this.BeginInvoke(new deleAppendLog(appendLog), "未检测到战网客户端");
                return;
            }
            if (psStatus[1])
            {
                if (!this.enableAutoRestart) {
                    return;
                }
                int status = this.getWowStatus();

                if (status == -1) {                    
                    this.BeginInvoke(new deleAppendLog(appendLog), "掉线，关闭魔兽世界");
                    this.killWow();
                } else if (status == 1) {
                    this.BeginInvoke(new deleAppendLog(appendLog), "人物选择界面，选择人物");
                    this.enterGame();                    
                }
            }
            else {
                this.BeginInvoke(new deleAppendLog(appendLog), "启动魔兽世界");
                this.launchWow();
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

        private int countMatchedPixels(LockBitmap lockbmp, int[,] pts) {

            Color st = Color.FromArgb(123, 9, 6);
            int ret = 0;

            int width = lockbmp.Width;
            int height = lockbmp.Height;
    
            for (int i = 0; i < pts.Length / 2; i++) {
                int ptx = pts[i,0] * width / 2560;
                int pty = pts[i,1] * height / 1440;
                if (calColorErr(st, lockbmp.GetPixel(ptx, pty)) < 500) {                   
                    ret += 1;
                }
            }
            return ret;
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

            IntPtr findPtr = this.activateWindow("魔兽世界");

            if (findPtr.ToInt32() != 0)
            {
                RECT cRect = new RECT();
                RECT wRect = new RECT();

                GetWindowRect(findPtr, ref wRect);
                GetClientRect(findPtr, ref cRect);

                Bitmap bmp = getScreenshot(ref cRect, ref wRect);
                LockBitmap lockbmp = new LockBitmap(bmp);
                lockbmp.LockBits();

                if (countMatchedPixels(lockbmp, loginPts) >= 10)
                {
                    return -1;
                }
                if (countMatchedPixels(lockbmp, choosePts) >= 10)
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
            IntPtr findPtr = FindWindow(null, "魔兽世界");
            if (findPtr.ToInt32() != 0) {
                RECT cRect = new RECT();
                RECT wRect = new RECT();
                GetClientRect(findPtr, ref cRect);
                GetWindowRect(findPtr, ref wRect);

                Point startPos = getClientStart(ref cRect, ref wRect);

                int cx = startPos.X + cRect.Right / 2;
                int cy = startPos.Y + cRect.Bottom * 1320 / 1440;
                SetCursorPos(cx, cy);
                mouse_event(0x0002, 0, 0, 0, 0);
                Thread.Sleep(90);
                mouse_event(0x0004, 0, 0, 0, 0);
            }
        }

        private IntPtr activateWindow(string title) {
            IntPtr findPtr = FindWindow(null, title);

            if (findPtr.ToInt32() != 0) {
                ShowWindow(findPtr, 9);
                SetForegroundWindow(findPtr);
            }
            return findPtr;
        }

        private void launchWow() {
            if (this.activateWindow("暴雪战网").ToInt32() != 0)
            {
                Thread.Sleep(300);
                SendKeys.SendWait("{ENTER}");
            }
            else {
                Console.WriteLine("BN NOT FOUND");
            }
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
            MessageBox.Show("使用说明\n\n1. 保持战网开启并选中魔兽世界\n2. 目前仅适配 16:9 屏幕，其它比例的显示器请使用 16:9 窗口模式\n\nBY: 小脑斧 2019/11/4");
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/Nihiue/LetMeRaid");
        }
        void textBox1_MouseDown(object sender, MouseEventArgs e)
        {
            HideCaret(((TextBox)sender).Handle);

        }

        private void textBox1_changed(object sender, EventArgs e)
        {
            HideCaret(((TextBox)sender).Handle);
        }
    }
}
