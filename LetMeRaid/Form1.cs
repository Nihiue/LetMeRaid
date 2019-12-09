﻿using System;
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
        private bool scheduleMode = false;
        private bool autoStartService = false;
        private bool ensureFocusBNWow = false;
        delegate void deleAppendLog(string text);
        [DllImport("kernel32.dll")]
        static extern uint SetThreadExecutionState(uint esFlags);
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

        private void preventSystemSleep(bool flag) {
            const uint ES_SYSTEM_REQUIRED = 0x00000001;
            const uint ES_DISPLAY_REQUIRED = 0x00000002;
            const uint ES_CONTINUOUS = 0x80000000;
            if (flag)
            {
                SetThreadExecutionState(ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED);
            }
            else {
                SetThreadExecutionState(ES_CONTINUOUS);
            }
        }

        bool isRatioSupported(int w, int h) {
            double r = (double)w / h;
            return r > 1.772;
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
            System.Timers.Timer timer = new System.Timers.Timer(8000);
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.onTick);
            timer.AutoReset = true;
            timer.Enabled = false;
            this.tickTimer = timer;
            if (this.autoStartService) {
                this.startService();
            }
        }

        public void onTick(object source, System.Timers.ElapsedEventArgs e) {

            if (this.scheduleMode) {
                DateTime dt = DateTime.Now;
                bool inTimeRange = dt.TimeOfDay.TotalMinutes >= (double)(this.numericUpDown1.Value * 60 + this.numericUpDown2.Value);

                if (!inTimeRange)
                {
                    return;
                }
            }
           
            bool[] psStatus = this.checkProcess();

            if (!psStatus[0]) {
                this.appendLog("未检测到战网客户端");
                return;
            }

            if (psStatus[1])
            {
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

        private int countMatchedPixels(LockBitmap lockbmp, int[,] pts) {

            Color st = Color.FromArgb(123, 9, 6);
            int ret = 0;

            int width = lockbmp.Width;
            int height = lockbmp.Height;

            int offsetX = (width - height * 16 / 9) / 2;

            for (int i = 0; i < pts.Length / 2; i++) {
                int ptx = offsetX + pts[i,0] * height / 2560 * 16 / 9;
                int pty = pts[i,1] * height / 1440;
                if (calColorErr(st, lockbmp.GetPixel(ptx, pty)) < 1600) {
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

            IntPtr findPtr = this.activateWindow("GxWindowClass", "魔兽世界");

            if (findPtr.ToInt32() != 0)
            {
                RECT cRect = new RECT();
                RECT wRect = new RECT();

                GetWindowRect(findPtr, ref wRect);
                GetClientRect(findPtr, ref cRect);

                if (!isRatioSupported(cRect.Right, cRect.Bottom)) {
                   this.appendLog("自动重设窗口大小");
                   MoveWindow(findPtr, 50, 50, 1280 + wRect.Right - wRect.Left - cRect.Right, 720 + wRect.Bottom - wRect.Top - cRect.Bottom, true);
                   return 0;
                }

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

        private bool ensureBNOnline() {
            IntPtr findPtr = this.activateWindow("Qt5QWindowIcon", "暴雪战网错误");
            if (findPtr.ToInt32() != 0) {
                this.appendLog("战网离线，尝试重连");
                SendKeys.SendWait("{ENTER}");
                return false;
            }
            return true;
        }

        private void launchWow() {
            IntPtr findPtr = this.activateWindow("Qt5QWindowOwnDCIcon", "暴雪战网");
            if (findPtr.ToInt32() != 0)
            {
                if (!this.ensureBNOnline()) {
                    return;
                }

                if (this.ensureFocusBNWow) {
                    MoveWindow(findPtr, 50, 50, 1380, 850, true);
                    Thread.Sleep(1000);
                    this.mouseClick(50 + 135, 50 + 155);
                    Thread.Sleep(1000);
                }
                // this.mouseClick(50 + 480, 50 + 750);

                SendKeys.SendWait("{ENTER}");
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
            if (!running)
            {
                this.updateDateControlStatus();
            }
            else {
                this.numericUpDown1.Enabled = !running;
                this.numericUpDown2.Enabled = !running;
            }

        }
        private void startService() {
            this.toggleUI(true);
            this.scheduleMode = !this.radioButton1.Checked;
            this.tickTimer.Enabled = true;
            this.preventSystemSleep(true);
            appendLog("启用服务");
        }
        private void stopService() {
            this.toggleUI(false);
            this.tickTimer.Enabled = false;
            this.preventSystemSleep(false);
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
            System.Diagnostics.Process.Start("https://www.nihi.me/lmr/");

            /*           string[] info = {
                "1. 保持战网开启并选中魔兽世界",
                "2. 适配 16:9 和 21:9 屏幕，其它比例会自动重设窗口大小",
                "3. 暂不支持多开",
                "",
                "BY: 小脑斧 2019/11/4"
            };
            MessageBox.Show(String.Join("\n", info), "使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
            */
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

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            this.updateDateControlStatus();
        }

        private void updateDateControlStatus() {
           this.numericUpDown1.Enabled = !this.radioButton1.Checked;
           this.numericUpDown2.Enabled = !this.radioButton1.Checked;
        }
    }
}
