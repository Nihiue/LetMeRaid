using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net;
using System.Collections.Specialized;
using System.IO;
using System.Net.Http;
using System.Text;

namespace LetMeRaid
{
    public partial class MainWindow : Form

    {
        private System.Timers.Timer tickTimer;
        private bool scheduleMode = false;
        private bool autoStartService = false;
        private bool focusFirstBNGame = false;
        private bool enableDebugLog = false;

        // Remote Report
        private string remoteReportToken = "";
        private bool remoteReportImage = false;
        private long lastRemoteReportTime = 0;
        private Bitmap lastScreenshot = null;

        private long asCounter = 0;
        private bool asFirstJoin = true;
        private int asHoldActionCnt = 0;
        private DateTime asStartTime = DateTime.Now;
        private bool castSpell = false;

        private FileStream debugFileStream;


        delegate void deleAppendLog(string text);
        [DllImport("kernel32.dll")]
        static extern uint SetThreadExecutionState(uint esFlags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
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
            public override string ToString()
            {
                return "{" + string.Format("Left={0},Top={1},Right={2},Bottom={3}", Left, Top, Right, Bottom) + "}";
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, System.Text.StringBuilder lpReturnedString, int nSize, string lpFileName);

        /*
        [DllImport("user32")]
        public static extern int GetSystemMetrics(int nIndex);
        */
        [DllImport("user32", EntryPoint = "HideCaret")]
        private static extern bool HideCaret(IntPtr hWnd);

        private string readConfigFile(string key, string def) {
            string filePath = System.IO.Path.Combine(Application.StartupPath, "config.ini");
            System.Text.StringBuilder sb = new System.Text.StringBuilder(1024);
            GetPrivateProfileString("LetMeRaid", key, def, sb, 1024, filePath);
            return sb.ToString();
        }
        private void loadConfig() {
            this.focusFirstBNGame = readConfigFile("FocusFirstBattleNetGame", "0") == "1";
            string t = readConfigFile("RemoteReportToken", "");
            if (t.Contains("@")) {
                this.remoteReportToken = t;
                this.appendLog("已配置远程日志");
            }

            this.remoteReportImage = readConfigFile("RemoteReportImage", "0") == "1";
            this.enableDebugLog = readConfigFile("DebugLog", "0") == "1";

            if (this.enableDebugLog) {
                this.appendLog("启用Debug日志输出");
            }
        }
        private void initDebugLog() {
            if (!this.enableDebugLog) {
                return;
            }
            string filePath = System.IO.Path.Combine(Application.StartupPath, "lmr_log.txt");
            this.debugFileStream = new FileStream(filePath, FileMode.Append);
            this.appendDebugLog("app start");
        }
        private static Point getClientStart(ref RECT cRect, ref RECT wRect) {
            Point start = new Point(wRect.Left, wRect.Top);
            // Console.WriteLine("({0},{1})  ({2},{3}) {4}", wRect.Right-wRect.Left, wRect.Bottom - wRect.Top, cRect.Right, cRect.Bottom, GetSystemMetrics(4));
            if (wRect.Bottom - wRect.Top != cRect.Bottom)
            {
                // 非全屏
                // int titleHeight = GetSystemMetrics(4);
                int borderWidth = (wRect.Right - wRect.Left - cRect.Right) / 2;
                start.X = start.X + borderWidth;
                start.Y = wRect.Bottom - cRect.Bottom - borderWidth;
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
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.loadConfig();
            this.initDebugLog();
            System.Timers.Timer timer = new System.Timers.Timer(1000);
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.onTick);
            timer.AutoReset = true;
            timer.Enabled = false;
            this.tickTimer = timer;
            if (this.autoStartService) {
                this.startService();
            }
        }

        public void onTick(object source, System.Timers.ElapsedEventArgs e) {

            if (this.tickTimer.Interval != 2 * 1000) {
                this.tickTimer.Interval = 2 * 1000;
            }

            if (this.scheduleMode) {
                DateTime dt = DateTime.Now;
                bool inTimeRange = dt.TimeOfDay.TotalMinutes >= (double)(this.numericUpDown1.Value * 60 + this.numericUpDown2.Value);

                if (!inTimeRange)
                {
                    return;
                }
            }

            bool[] psStatus = this.checkProcess();

            if (!psStatus[0])
            {
                this.appendLog("未检测到战网客户端");
            }
            else {
                if (psStatus[1])
                {
                    if (psStatus[2])
                    {
                        // 检测到 teamviewer， 关闭弹窗以防止影响截图
                        this.closeTeamviewerPopup();
                    }
                    int status = this.getWowStatus();

                    if (status == -1)
                    {
                        this.appendLog("掉线，关闭魔兽世界");
                        this.killWow();
                    }
                    else if (status == 1)
                    {
                        this.appendLog("人物选择界面，选择人物");
                        this.enterGame();
                    }
                    else
                    {
                        // 自动奥山
                        this.autoAS();
                    }
                }
                else
                {
                    this.appendLog("启动魔兽世界");
                    this.launchWow();
                }
            }

            if (this.enableDebugLog) {
                this.debugFileStream.Flush();
                if (this.lastScreenshot != null) {
                    this.writeCapImage(this.lastScreenshot);
                }
            }

            long ts = (new DateTimeOffset(DateTime.UtcNow)).ToUnixTimeSeconds();
            if (this.remoteReportToken.Contains("@") && ts - this.lastRemoteReportTime > 298) {
                this.lastRemoteReportTime = ts;
                this.sendReportAsync(String.Join("\n", this.textBox1.Lines), this.lastScreenshot);
            }

            if (this.lastScreenshot != null) {
                this.lastScreenshot.Dispose();
                this.lastScreenshot = null;
            }
        }

        private void writeCapImage(Bitmap bmp) {
            string filePath = System.IO.Path.Combine(Application.StartupPath, "lmr_cap.jpg");
            var stream = new FileStream(filePath, FileMode.Create);
            bmp.Save(stream, ImageFormat.Jpeg);
        }

        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        private byte[] encodeImage(Bitmap bmp) {
            int resHeight = 720;
            Bitmap resizedImage = new Bitmap(bmp, new Size(bmp.Width * resHeight / bmp.Height, resHeight));

            ImageCodecInfo jgpEncoder = GetEncoder(ImageFormat.Jpeg);
            System.Drawing.Imaging.Encoder myEncoder = System.Drawing.Imaging.Encoder.Quality;

            var myEncoderParameters = new EncoderParameters(1);
            var myEncoderParameter = new EncoderParameter(myEncoder, 60L);
            myEncoderParameters.Param[0] = myEncoderParameter;

            var stream = new MemoryStream();
            resizedImage.Save(stream, jgpEncoder, myEncoderParameters);

            resizedImage.Dispose();
            byte[] imageBytes = stream.ToArray();
            return imageBytes;
        }

        private async void sendReportAsync(string log, Bitmap bmp) {
            try
            {
                HttpClient httpClient = new HttpClient();
                MultipartFormDataContent form = new MultipartFormDataContent();
                form.Add(new StringContent(this.remoteReportToken), "token");
                form.Add(new StringContent(log), "log");
                if (this.remoteReportImage && bmp != null)
                {
                    byte[] file_bytes = this.encodeImage(bmp);
                    form.Add(new ByteArrayContent(file_bytes, 0, file_bytes.Length), "image", "screenshot.jpg");
                }
                await httpClient.PostAsync("http://ap.nihi.me/lmr_api/report", form);
            }
            catch (Exception e)
            {
                this.appendDebugLog("remote report failed");
                Console.WriteLine(e);
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

        private int countMatchedPixels(Bitmap bmp, int[,] pts) {

            Color st = Color.FromArgb(123, 9, 6);
            int ret = 0;

            int width = bmp.Width;
            int height = bmp.Height;

            int offsetX = (width - height * 16 / 9) / 2;

            for (int i = 0; i < pts.Length / 2; i++) {
                int ptx = offsetX + pts[i,0] * height / 2560 * 16 / 9;
                int pty = pts[i,1] * height / 1440;
                if (calColorErr(st, bmp.GetPixel(ptx, pty)) < 1600) {
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

                this.appendDebugLog(string.Format("cRect={0} wRect={1}", cRect, wRect));

                if (!isRatioSupported(cRect.Right, cRect.Bottom)) {
                   int newWidth = 1280 + wRect.Right - wRect.Left - cRect.Right;
                   int newHeight = 720 + wRect.Bottom - wRect.Top - cRect.Bottom;
                   this.appendLog(string.Format("重设窗口大小 {0}*{1}", newWidth, newHeight));
                   MoveWindow(findPtr, 50, 50, newWidth, newHeight, true);
                   return 0;
                }

                Bitmap bmp = getScreenshot(ref cRect, ref wRect);

                int ret = 0;
                int loginMatched = countMatchedPixels(bmp, loginPts);
                int chooseMatched = countMatchedPixels(bmp, choosePts);

                this.appendDebugLog(string.Format("pixel count login={0} choose={1}", loginMatched, chooseMatched));
                if (loginMatched >= 10)
                {
                    ret = -1;
                }
                else if (chooseMatched >= 10)
                {
                    ret = 1;
                }
                this.lastScreenshot = bmp;
                return ret;
            }
            else
            {
                this.appendLog("未找到 WOW 窗口");
            }
            return 0;
        }

        private static Bitmap cropAtRect(Bitmap b, Rectangle r)
        {
            Bitmap nb = new Bitmap(r.Width, r.Height);
            Graphics g = Graphics.FromImage(nb);
            g.DrawImage(b, -r.X, -r.Y);
            return nb;
        }

        private void resetASVars()
        {
            this.asCounter = 0;
            this.asHoldActionCnt = 0;
        }

        private void autoAS() {
            // 0: 目标+排队奥山宏
            // 9: 确认加入宏
            // 8: 徽章
            // 7: afk宏

            // 截图进行OCR
            Bitmap bmp = lastScreenshot;
            double ratio = bmp.Width / 3840.0;
            Console.WriteLine(ratio);
            Bitmap nb = cropAtRect(bmp, new Rectangle((int)(20 * ratio), (int)(60 * ratio), (int)(800 * ratio), (int)(85 * ratio)));
            string filePath = System.IO.Path.Combine(Application.StartupPath, "test.jpg");
            var stream = new FileStream(filePath, FileMode.Create);
            nb.Save(stream, ImageFormat.Jpeg);
            stream.Close();
            MemoryStream ms = new MemoryStream();
            nb.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            byte[] bytes = ms.GetBuffer();
            ms.Close();
            var engine = new Tesseract.TesseractEngine(@"./tessdata", "eng", Tesseract.EngineMode.Default);
            var page = engine.Process(Tesseract.Pix.LoadFromMemory(bytes));
            var text = page.GetText().Trim();
            Console.WriteLine(text);

            // 获取战场状态与人物状态
            var status = "QUEUED";
            var aliveStatus = "ALIVE";
            try
            {
                var lines = text.Split(',');
                status = lines[0].Split(':')[1].Trim();
                aliveStatus = lines[1].Split(':')[1].Trim();
            } catch {
                appendLog(string.Format("识别失败, 原始数据：\n{0}", text));
            }

            // var zone = lines[1].Split(':')[0];
            //var locX = lines[1].Split(':')[1].Split(',')[0];
            //var locY = lines[1].Split(':')[1].Split(',')[1];
            //Console.WriteLine("{0}, {1}, {2}, {3}", status, zone, locX, locY);
            Console.WriteLine(status);
            Console.WriteLine(aliveStatus);
            appendLog(string.Format("status: {0}, aliveStatus:{1}", status, aliveStatus));

            var input = new WindowsInput.InputSimulator();
            if (!(status == "ACTIVE" && aliveStatus == "ALIVE" && asHoldActionCnt > 0))
            {
                // 除了持续型操作，都释放前进按键
                // input.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_W);
            }

            // 根据不同战场状态进行操作
            if (status == "NONE")
            {
                input.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_W);
                // 在战场外未排队，进行自动排队，初始化各标记值
                resetASVars();
                asFirstJoin = true;

                SendKeys.SendWait("0");
                Thread.Sleep(300);
                SendKeys.SendWait("j");
                Thread.Sleep(300);
                SendKeys.SendWait("0");
                Thread.Sleep(300);
                SendKeys.SendWait("0");
            } else if (status == "QUEUED")
            {
                // 排队中，跳跃
                if (DateTime.Now.Second % 9 == 0)
                {
                    SendKeys.SendWait(" ");
                }
            }
            else if (status == "CONFIRM" || status == "CONFIRN")
            {
                // 确认进入战场
                Thread.Sleep(1000);
                SendKeys.SendWait("9");
                asStartTime = DateTime.Now;
            }
            else if (status == "ACTIVE")
            {
                // 战场内自动操作
                // 超过一小时退出
                var now = DateTime.Now;
                var duration = now - asStartTime;
                if (duration.TotalHours >= 1)
                {
                    input.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_W);
                    SendKeys.SendWait("7");
                    return;
                }

                // 判断人物当前状态
                if (aliveStatus == "DEAD" || aliveStatus == "GHOST")
                {
                    // 死亡或者幽灵，初始化标记值，停止行动等待复活
                    input.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_W);

                    resetASVars();
                    asFirstJoin = false;

                    if (aliveStatus == "DEAD")
                    {
                        // 自动释放灵魂
                        SendKeys.SendWait("9");
                    }
                } else if (aliveStatus == "ALIVE")
                {
                    // 正常状态，执行自动行动，增加计数器
                    asCounter += 1;

                    // 处理持续多个 Tick 的操作
                    if (asHoldActionCnt > 0)
                    {
                        asHoldActionCnt -= 1;
                        if (!castSpell)
                        {
                            SendKeys.SendWait(" ");
                        }
                        return;
                    }
                   
                    // 使用徽章的轮次
                    var badgeInterval = 70;

                    // 开始正常行动的轮次
                    // 刚加入时需要等待战场开门
                    var normalActionInterval = asFirstJoin ? 60 : 20;

                    Random rnd = new Random();

                    if (asCounter > normalActionInterval)
                    {
                        if (asFirstJoin)
                        {
                            // 在往前走30s
                            asFirstJoin = false;
                            // input.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.RIGHT);
                            // Thread.Sleep(30);
                            // input.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.RIGHT);
                            input.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_W);
                            asHoldActionCnt = 25;
                            return;
                        }
                        // 正常行动
                        // 判断是否要使用徽章
                        this.castSpell = false;
                        if (asCounter % badgeInterval == 0)
                        {
                            // Thread.Sleep(1000);
                            // SendKeys.SendWait("8");
                            // 徽章持续10s，所以下个tick不动
                            // asHoldActionCnt = 1;
                            // return;
                        }

                        // 随机攻击
                        SendKeys.SendWait("{TAB}");
                        var castSpell = rnd.Next(0, 3);
                        if (castSpell == 0)
                        {
                            SendKeys.SendWait("f");
                        } else if (castSpell == 1)
                        {
                            SendKeys.SendWait("e");
                        }

                        // 随机转头
                        var keyCode = WindowsInput.Native.VirtualKeyCode.LEFT;
                        if (rnd.Next(0, 2) == 1)
                        {
                            keyCode = WindowsInput.Native.VirtualKeyCode.RIGHT;
                        }
                        input.Keyboard.KeyDown(keyCode);
                        Thread.Sleep(rnd.Next(0, 700));
                        input.Keyboard.KeyUp(keyCode);

                    }
                    else if (asCounter == 1)
                    {
                        // 前两个tick，精确处理山洞路线
                        var keyCode = WindowsInput.Native.VirtualKeyCode.LEFT;

                        input.Keyboard.KeyDown(keyCode);
                        Thread.Sleep(105);
                        input.Keyboard.KeyUp(keyCode);
                        input.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_W);
                        asHoldActionCnt = 2;
                        return;

                    }
                    else if (asCounter == 4)
                    {
                        var keyCode = WindowsInput.Native.VirtualKeyCode.RIGHT;

                        input.Keyboard.KeyDown(keyCode);
                        Thread.Sleep(65);
                        input.Keyboard.KeyUp(keyCode);
                        input.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_W);
                        asHoldActionCnt = 3;
                        return;
                    }
                    else if (asFirstJoin)
                    {
                        // 出门前停止行动
                        input.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_W);
                        return;
                    }

                    // 自动前进，夹杂跳跃
                    input.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_W);
                    if (asCounter % 3 == 0)
                    {
                        SendKeys.SendWait(" ");
                    }
                }
            }

            //System.Environment.Exit(0);
        }

        private void killWow() {
            foreach (Process p in Process.GetProcesses())
            {
                if (p.ProcessName.ToLower() == "wow" || p.ProcessName.ToLower() == "wowclassic")
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
                ShowWindow(findPtr, 5);
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

                if (this.focusFirstBNGame) {
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
                if (p.ProcessName.ToLower() == "wow" || p.ProcessName.ToLower() == "wowclassic")
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

        private void appendDebugLog(string log) {
            if (!this.enableDebugLog) {
                return;
            }
            try
            {
                DateTime dt = DateTime.Now;
                string line = string.Format("{0:G} {1}\n", dt, log);
                byte[] byets = System.Text.Encoding.UTF8.GetBytes(line);
                this.debugFileStream.Write(byets, 0, byets.Length);
            }
            catch (Exception ex) {
               Debug.WriteLine(string.Format("写入文件出错：消息={0},堆栈={1}", ex.Message, ex.StackTrace));
            }
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
            this.tickTimer.Interval = 1000;
            this.tickTimer.Enabled = true;
            this.preventSystemSleep(true);
            appendLog("启用服务");
            this.appendDebugLog("service start");
        }
        private void stopService() {
            this.toggleUI(false);
            this.tickTimer.Enabled = false;
            this.preventSystemSleep(false);
            appendLog("停止服务");
            this.appendDebugLog("service stop");
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
