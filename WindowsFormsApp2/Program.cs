using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Runtime.InteropServices;
//using System.Threading;
using System.Diagnostics;

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

            Application.Run(new CustomApplicationContext());
        }
    }

    public class CustomApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        static AutomationElement desktop = AutomationElement.RootElement;
        static AutomationElement tasklist = desktop.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "MSTaskListWClass"));
        static AutomationElement parent = TreeWalker.ControlViewWalker.GetParent(tasklist);
        //static AutomationElement taskbar = desktop.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "ReBarWindow32"));

        private IntPtr tasklistPtr;

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,
        }

        public CustomApplicationContext()
        {
            MenuItem header = new MenuItem("CenterTaskbar", Pass);
            header.Enabled = false;

            // Setup Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.Icon1,
                ContextMenu = new ContextMenu(new MenuItem[] {
                header,
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };

            tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

            if (SafetyCheck()) {
                Debug.WriteLine("Hooked Automation Class Targets:");
                Debug.WriteLine(desktop.Current.ClassName);
                Debug.WriteLine(tasklist.Current.ClassName);
                Debug.WriteLine(parent.Current.ClassName);
            }

            Reposition();
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, TreeScope.Subtree, OnAutomation);
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, TreeScope.Subtree, OnAutomation);
        }

        void Pass(object sender, EventArgs e)
        {
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;
            Reset();
            System.Windows.Forms.Application.Exit();
        }

        private void OnAutomation(object sender, AutomationEventArgs e)
        {
            // Windows changed, update position
            Reposition();
        }

        private bool SafetyCheck()
        {
            if (desktop == null || tasklist == null || parent == null)
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

            return ScreenScalingFactor;
        }

        /*
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
        */

        private void Reset()
        {
            if (!SafetyCheck())
            {
                return;
            }

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect barBounds = parent.Current.BoundingRectangle;
            Double deltax = Math.Abs(listBounds.X - barBounds.X);

            if (deltax <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call. (DeltaX = " + deltax + ")");
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
            Debug.WriteLine("Begin Reposition Calculation");
            if (!SafetyCheck())
            {
                Debug.WriteLine("Failed safety check, aborting");
                return;
            }

            AutomationElementCollection children = tasklist.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
            if (children == null || children.Count < 1)
            {
                Debug.WriteLine("Failed to find any icons, aborting");
                return;
            }

            int screenWidth = Screen.PrimaryScreen.Bounds.Width;
            double scalingFactor = getScalingFactor();

            AutomationElement first = children[0];
            AutomationElement last = children[children.Count - 1];
            Double width = last.Current.BoundingRectangle.Right - first.Current.BoundingRectangle.Left;
            width = width / getScalingFactor();

            double targetX = Math.Round((screenWidth - width) / 2);
            Debug.Write("Desktop width: ");
            Debug.WriteLine(desktop.Current.BoundingRectangle.Width);
            Debug.Write("Screen width: ");
            Debug.WriteLine(screenWidth);
            Debug.Write("Total icon width: ");
            Debug.WriteLine(width);
            Debug.Write("Target abs X position: ");
            Debug.WriteLine(targetX);
            Debug.Write("Calculated Total Width: ");
            Debug.WriteLine(Math.Round(targetX + width + targetX));

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect parentBounds = parent.Current.BoundingRectangle;
            Double deltax = Math.Abs(targetX - listBounds.X);

            // Previous bounds check
            if (deltax <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call (DeltaX = " + deltax + ")");
                return; 
            }

            // Right bounds check
            double safeRight = screenWidth - parentBounds.Right;
            if ((targetX + width) > (screenWidth - safeRight))
            {
                // Shift off center when the bar is too big
                Debug.WriteLine("Shifting off center, too big and hitting right boundary (" + (targetX + width) + " > " + (screenWidth - safeRight) + ")");
                double extra = (targetX + listBounds.Width) - (screenWidth - safeRight);
                targetX -= extra;
            }

            // Left bounds check
            if (targetX <= (parentBounds.X))
            {
                // Prevent X position ending up beyond the normal left aligned position
                Debug.WriteLine("Target is more left than left aligned default, left aligning (" + targetX + " <= " + parentBounds.X + ")");
                Reset();
            }

            int oldWidth = (int)listBounds.Width;
            int oldHeight = (int)listBounds.Height;
 
            MoveWindow(tasklistPtr, relativeX(targetX), 0, oldWidth, oldHeight, true);

            Debug.Write("Final X Position: ");
            Debug.WriteLine(tasklist.Current.BoundingRectangle.X);
            Debug.WriteLine((tasklist.Current.BoundingRectangle.X == targetX) ? "Move hit target": "Move missed target");
        }

        private int relativeX(double x)
        {
            Rect barBounds = parent.Current.BoundingRectangle;
            Double newPos = x - barBounds.X;

            if (newPos < 0)
            {
                Debug.WriteLine("Relative position < 0, adjusting to 0 (Previous: " + newPos + ")");
                newPos = 0;
            }

            return (int)newPos;
        }
    }
}
