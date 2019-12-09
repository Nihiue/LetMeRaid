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
        private bool autoStartService = false;
        private bool ensureFocusBNWow = false;
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
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32")]
        public static extern int GetSystemMetrics(int nIndex);
        [DllImport("user32", EntryPoint = "HideCaret")]
        private static extern bool HideCaret(IntPtr hWnd);
        private static Point getClientStart(ref RECT cRect, ref RECT wRect) {
            Point start = new Point(wRect.Left, wRect.Top);
            // Console.WriteLine("({0},{1})  ({2},{3}) {4}", wRect.Right-wRect.Left, wRect.Bottom - wRect.Top, cRect.Right, cRect.Bottom, GetSystemMetrics(4));
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

        private int checkRatioType(int w, int h) {
            double r = (double)w / h;
            if (r > 1.772 && r < 1.783) {
                // 16:9
                return 1;
            }
            if (r > 2.37 && r < 2.39) {
                // 21:9
                return 2;
            }
            return 0;
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

        public MainWindow(string[] args)
        {
            InitializeComponent();

            if (Array.IndexOf(args, "--auto-start") >= 0) {
                this.autoStartService = true;
                this.ensureFocusBNWow = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Timers.Timer timer = new System.Timers.Timer(5000);
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.onTick);
            timer.AutoReset = true;
            timer.Enabled = false;
            this.tickTimer = timer;
            if (this.autoStartService) {
                this.startService();
            }
        }

        public void onTick(object source, System.Timers.ElapsedEventArgs e) {
            DateTime dt = DateTime.Now;
            bool inTimeRange = dt.TimeOfDay.TotalMinutes >= (double)(this.numericUpDown1.Value * 60 + this.numericUpDown2.Value);

            if (!inTimeRange) {
                return;
            }

            bool[] psStatus = this.checkProcess();

            if (!psStatus[0]) {
                this.appendLog("未检测到战网客户端");
                return;
            }

            if (psStatus[1])
            {
                if (!this.enableAutoRestart) {
                    return;
                }

                if (psStatus[2])
                {
                    // 检测到 teamviewer， 关闭弹窗以防止影响截图
                    this.closeTeamviewerPopup();
                }
                int status = this.getWowStatus();

                if (status == -1) {
                    this.appendLog("掉线，关闭魔兽世界");
                    this.killWow();
                } else if (status == 1) {
                    this.appendLog("人物选择界面，选择人物");
                    this.enterGame();
                }
            }
            else {
                this.appendLog("启动魔兽世界");
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

        private int countMatchedPixels(LockBitmap lockbmp, int[,] pts, int unitWidth, int unitHeight) {

            Color st = Color.FromArgb(123, 9, 6);
            int ret = 0;

            int width = lockbmp.Width;
            int height = lockbmp.Height;

            for (int i = 0; i < pts.Length / 2; i++) {
                int ptx = pts[i,0] * width / unitWidth;
                int pty = pts[i,1] * height / unitHeight;
                if (calColorErr(st, lockbmp.GetPixel(ptx, pty)) < 1200) {
                    ret += 1;
                }
            }
            return ret;
        }
        private int getWowStatus() {
            int[,] loginPts_16_9 = {
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

            int[,] choosePts_16_9 = {
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

            int[,] loginPts_21_9 = {
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

            int[,] choosePts_21_9 = {
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

            IntPtr findPtr = this.activateWindow("GxWindowClass", "魔兽世界");

            if (findPtr.ToInt32() != 0)
            {
                RECT cRect = new RECT();
                RECT wRect = new RECT();

                GetWindowRect(findPtr, ref wRect);
                GetClientRect(findPtr, ref cRect);

                int ratioType = this.checkRatioType(cRect.Right, cRect.Bottom);

                if (ratioType == 0) {
                    if (wRect.Bottom - wRect.Top != cRect.Bottom)
                    {
                        // 窗口模式
                        this.appendLog("自动重设窗口大小");
                        MoveWindow(findPtr, 50, 50, 1280 + wRect.Right - wRect.Left - cRect.Right, 720 + wRect.Bottom - wRect.Top - cRect.Bottom, true);
                    }
                    else {
                        this.appendLog("不支持当前屏幕比例, 请启用窗口模式");
                    }
                    return 0;
                }

                Bitmap bmp = getScreenshot(ref cRect, ref wRect);
                LockBitmap lockbmp = new LockBitmap(bmp);
                lockbmp.LockBits();

                if (ratioType == 1) {
                    // 16:9
                    if (countMatchedPixels(lockbmp, loginPts_16_9, 2560, 1440) >= 10)
                    {
                        return -1;
                    }
                    if (countMatchedPixels(lockbmp, choosePts_16_9, 2560, 1440) >= 10)
                    {
                        return 1;
                    }

                } else if (ratioType == 2) {
                    // 21:9
                    if (countMatchedPixels(lockbmp, loginPts_21_9， 2560, 1080) >= 10)
                    {
                        return -1;
                    }
                    if (countMatchedPixels(lockbmp, choosePts_21_9, 2560, 1080) >= 10)
                    {
                        return 1;
                    }
                }
            }
            else
            {
                this.appendLog("未找到 WOW 窗口");
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
            IntPtr findPtr = this.activateWindow("GxWindowClass", "魔兽世界");
            if (findPtr.ToInt32() != 0) {
                RECT cRect = new RECT();
                RECT wRect = new RECT();
                GetClientRect(findPtr, ref cRect);
                GetWindowRect(findPtr, ref wRect);

                Point startPos = getClientStart(ref cRect, ref wRect);

                int cx = startPos.X + cRect.Right / 2;
                int cy = startPos.Y + cRect.Bottom * 1320 / 1440;
                this.mouseClick(cx, cy, 85);
            }
        }

        private IntPtr activateWindow(string wClass, string title) {
            IntPtr findPtr = FindWindow(wClass, title);

            if (findPtr.ToInt32() != 0) {
                ShowWindow(findPtr, 9);
                SetForegroundWindow(findPtr);
                Thread.Sleep(300);
            }
            return findPtr;
        }

        private void mouseClick(int x, int y, int duration = 85) {
            SetCursorPos(x, y);
            mouse_event(0x0002, 0, 0, 0, 0);
            Thread.Sleep(duration);
            mouse_event(0x0004, 0, 0, 0, 0);
        }

        private void launchWow() {
            IntPtr findPtr = this.activateWindow("Qt5QWindowOwnDCIcon", "暴雪战网");
            if (findPtr.ToInt32() != 0)
            {
                MoveWindow(findPtr, 50, 50, 1380, 850, true);
                Thread.Sleep(1000);
                if (this.ensureFocusBNWow) {
                    this.mouseClick(50 + 135, 50 + 155);
                    Thread.Sleep(1000);
                }
                this.mouseClick(50 + 480, 50 + 750);
            }
            else {
                this.appendLog("未找到 Battle.net 窗口");
            }
        }

        private bool[] checkProcess() {
            bool[] ret = { false, false, false };
            Process[] ps = Process.GetProcesses();
            foreach (Process p in ps) {
                if (p.ProcessName.ToLower() == "battle.net") {
                    ret[0] = true;
                }
                if (p.ProcessName.ToLower() == "wow")
                {
                    ret[1] = true;
                }
                if (p.ProcessName.ToLower() == "teamviewer")
                {
                    ret[2] = true;
                }
            }
                return ret;
        }

        private void closeTeamviewerPopup() {
            IntPtr findPtr = this.activateWindow("#32770", "Sponsored session");
            if (findPtr.ToInt32() == 0) {
                findPtr = this.activateWindow("#32770", "发起会话");
            }
            if (findPtr.ToInt32() != 0) {
                this.appendLog("关闭 TeamViewer 弹窗");
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(500);
            }
        }

        private void appendLog(string log) {
            DateTime dt = DateTime.Now;
            string line = string.Format("{0:T} {1}", dt, log);
            this.BeginInvoke(new deleAppendLog(updateLogBox), line);
        }

        private void updateLogBox(string line) {
            List<string> tmp = this.textBox1.Lines.ToList();
            tmp.Add(line);
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
            string[] info = {
                "1. 保持战网开启并选中魔兽世界",
                "2. 目前仅适配 16:9 屏幕，其它比例的显示器请使用窗口模式",
                "3. 暂不支持多开",
                "",
                "BY: 小脑斧 2019/11/4"
            };
            MessageBox.Show(String.Join("\n", info), "使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
