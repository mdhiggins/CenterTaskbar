using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Threading;

namespace CenterTaskbar
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool thread = (args.Length > 0);

            Application.Run(new CustomApplicationContext(thread));
        }
    }

    public class CustomApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        static AutomationElement desktop = AutomationElement.RootElement;
        static AutomationElement tasklist = desktop.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "MSTaskListWClass"));
        static AutomationElement taskbar = desktop.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "ReBarWindow32"));

        Double padding = 2; // 2px space between ReBarWindow32 and MSTaskListWClass apparently
        IntPtr tasklistPtr;
        bool running = false;

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,
        }

        public CustomApplicationContext(bool thread)
        {
            String title = "CenterTaskbar Listening";
            if (thread)
            {
                title = "CenterTaskbar Looping";
            }
            // Setup Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.Icon1,
                ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem(title, Pass),
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };

            tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

            if (thread) {
                Loop();
            } else {
                Reposition();
                Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, TreeScope.Subtree, OnAutomation);
                Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, TreeScope.Subtree, OnAutomation);
            }
        }

        void Pass(object sender, EventArgs e)
        {
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;
            running = false;
            Reset();
            System.Windows.Forms.Application.Exit();
        }

        private void OnAutomation(object sender, AutomationEventArgs e)
        {
            // Windows changed, update position
            Console.WriteLine("Property Changed");
            Reposition();
        }

        private bool SafetyCheck()
        {
            if (desktop == null || tasklist == null || taskbar == null)
            {
                return false;
            }
            return true;
        }

        private double getScalingFactor()
        {
            Graphics g = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();
            int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);

            double ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;

            return ScreenScalingFactor; // 1.25 = 125%
        }

        private void Loop()
        {
            new Thread(new ThreadStart(this.LoopThread)).Start();
        }

        private void LoopThread()
        {
            running = true;
            while (running)
            {

                Thread.Sleep(150);
                Application.DoEvents();
                Reposition();
            }
            Reset();
        }

        private void Reset()
        {
            if (!SafetyCheck())
            {
                return;
            }

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect barBounds = taskbar.Current.BoundingRectangle;
            Double deltax = listBounds.X - padding - barBounds.X;

            if (Math.Abs(deltax) <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Console.WriteLine("Already positioned, ending");
                return;
            }

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;
            Rect bounds = tasklist.Current.BoundingRectangle;
            int newWidth = (int)bounds.Width;
            int newHeight = (int)bounds.Height;
            MoveWindow(tasklistPtr, 0, 0, newWidth, newHeight, true);
        }

        private void Reposition()
        {
            if (!SafetyCheck())
            {
                return;
            }

            AutomationElementCollection children = tasklist.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
            if (children == null || children.Count < 1)
            {
                return;
            }

            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            double scalingFactor = getScalingFactor();

            AutomationElement first = children[0];
            AutomationElement last = children[children.Count - 1];
            Double width = last.Current.BoundingRectangle.Right - first.Current.BoundingRectangle.Left;
            width = width / getScalingFactor();

            double targetX = Math.Round((screenWidth / 2) - (width / 2));

            Console.Write("Screen width: ");
            Console.WriteLine(screenWidth);
            Console.Write("Total width: ");
            Console.WriteLine(width);
            Console.Write("Target X: ");
            Console.WriteLine(targetX);

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect barBounds = taskbar.Current.BoundingRectangle;
            Double deltax = targetX - (listBounds.X);

            //Console.WriteLine(deltax);

            if (Math.Abs(deltax) <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Console.WriteLine("Already positioned, ending");
                return; 
            }

            Console.Write("Right Bar Bounds: ");
            Console.WriteLine(barBounds.Right);
            double safeRight = screenWidth - barBounds.Right;
            Console.WriteLine(safeRight);
            if ((targetX + width) > (screenWidth - safeRight))
            {
                // Shift off center when the bar is too big
                Console.WriteLine("Shifting off center, too big");
                double extra = (targetX + listBounds.Width) - (screenWidth - safeRight);
                targetX -= extra;
            }

            if (targetX <= (barBounds.X + padding))
            {
                // Prevent X position ending up beyond the normal left aligned position
                Console.WriteLine("Target is more left than left aligned default, resetting");
                Reset();
            }

            int oldWidth = (int)listBounds.Width;
            int oldHeight = (int)listBounds.Height;
 
            MoveWindow(tasklistPtr, relativeX(targetX), 0, oldWidth, oldHeight, true);

            Console.Write("Final Position: ");
            Console.WriteLine(tasklist.Current.BoundingRectangle.X);
        }

        private int relativeX(double x)
        {
            Rect barBounds = taskbar.Current.BoundingRectangle;
            Double newPos = x - barBounds.X - padding; // MoveWindow uses relative positioning

            if (newPos < 0)
            {
                newPos = 0;
            }

            return (int)newPos;
        }
    }
}
