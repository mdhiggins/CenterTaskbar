using System;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;

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
        static String MSTaskListWClass = "MSTaskListWClass";
        //static String ReBarWindow32 = "ReBarWindow32";
        static String Shell_TrayWnd = "Shell_TrayWnd";
        static String Shell_SecondaryTrayWnd = "Shell_SecondaryTrayWnd";

        Dictionary<String, bool> running = new Dictionary<String, bool>();
        Dictionary<String, double> targets = new Dictionary<String, double>();

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
            MenuItem header = new MenuItem("CenterTaskbar", Exit);
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

            UpdateAll();

            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, TreeScope.Subtree, OnAutomation);
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, TreeScope.Subtree, OnAutomation);
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;
            UpdateAll(true);
            System.Windows.Forms.Application.Exit();
        }

        private void OnAutomation(object sender, AutomationEventArgs e)
        {
            // Windows changed, update position
            UpdateAll();
        }

        private bool SafetyCheck()
        {
            if (desktop == null)
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
            Debug.Write("Scaling factor: ");
            Debug.WriteLine(ScreenScalingFactor);
            return ScreenScalingFactor;
        }

        private void Reset(AutomationElement trayWnd)
        {
            if (!SafetyCheck())
            {
                return;
            }

            AutomationElement tasklist = trayWnd.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass));
            if (tasklist == null)
            {
                Debug.WriteLine("Null values found, aborting reset");
                return;
            }
            AutomationElement tasklistcontainer = TreeWalker.ControlViewWalker.GetParent(tasklist);
            if (tasklistcontainer == null)
            {
                Debug.WriteLine("Null values found, aborting reset");
                return;
            }
            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect barBounds = tasklistcontainer.Current.BoundingRectangle;
            Double deltax = Math.Abs(listBounds.X - barBounds.X);

            if (deltax <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call. (DeltaX = " + deltax + ")");
                return;
            }

            Rect bounds = tasklist.Current.BoundingRectangle;
            int newWidth = (int)bounds.Width;
            int newHeight = (int)bounds.Height;
            MoveWindow(tasklistPtr, 0, 0, newWidth, newHeight, true);
        }

        private AutomationElement GetTopLevelWindow(AutomationElement element)
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement elementParent;
            AutomationElement node = element;
            do
            {
                if (node == null)
                {
                    break;
                }
                elementParent = walker.GetParent(node);
                Debug.WriteLine(elementParent.Current.ClassName);
                Debug.WriteLine(elementParent.Current.BoundingRectangle.Width);
                if (elementParent == AutomationElement.RootElement) break;
                node = elementParent;
            }
            while (true);
            return node;
        }

        private void UpdateAll(bool reset = false)
        {
            OrCondition condition = new OrCondition(new PropertyCondition(AutomationElement.ClassNameProperty, Shell_TrayWnd), new PropertyCondition(AutomationElement.ClassNameProperty, Shell_SecondaryTrayWnd));
            AutomationElementCollection lists = desktop.FindAll(TreeScope.Children, condition);
            Debug.WriteLine(lists.Count + " bar(s) detected");
            foreach (AutomationElement trayWnd in lists)
            {
                new Thread(() =>
                {
                    Thread.CurrentThread.IsBackground = true;
                    Thread.Sleep(400); // Delay which gives times for resizing to finish
                    if (reset)
                    {
                        Reset(trayWnd);
                    }
                    else
                    {
                        Reposition(trayWnd);
                    }
                }).Start();
            }
        }

        private void Reposition(AutomationElement trayWnd)
        {
            Debug.WriteLine("Begin Reposition Calculation");
            AutomationElement tasklist = trayWnd.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass));
            if (tasklist == null)
            {
                Debug.WriteLine("Null values found, aborting");
                return;
            }
            AutomationElement tasklistcontainer = TreeWalker.ControlViewWalker.GetParent(tasklist);
            if (tasklistcontainer == null)
            {
                Debug.WriteLine("Null values found, aborting");
                return;
            }

            Rect tasklistBounds = tasklist.Current.BoundingRectangle;
            Rect tasklistContainerBounds = tasklistcontainer.Current.BoundingRectangle;
            Rect trayBounds = trayWnd.Current.BoundingRectangle;

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;
            IntPtr trayWndPtr = (IntPtr)trayWnd.Current.NativeWindowHandle;

            Double barWidth = trayWnd.Current.BoundingRectangle.Width;
            Double scalingFactor = getScalingFactor();

            AutomationElement first = TreeWalker.ControlViewWalker.GetFirstChild(tasklist);
            AutomationElement last = TreeWalker.ControlViewWalker.GetLastChild(tasklist);

            if (first == null || last == null)
            {
                Debug.WriteLine("Null values found for items, aborting");
                return;
            }

            Double width = (last.Current.BoundingRectangle.Right - first.Current.BoundingRectangle.Left) / scalingFactor;
            Double xOffset = tasklistContainerBounds.X;
            Double targetX = Math.Round((barWidth - width) / 2) + trayBounds.X;

            Debug.Write("Bar width: ");
            Debug.WriteLine(barWidth);
            Debug.Write("Total icon width: ");
            Debug.WriteLine(width);
            Debug.Write("Target abs X position: ");
            Debug.WriteLine(targetX);
            Debug.Write("Reconstructed Width: ");
            Debug.WriteLine(Math.Round(targetX + width + targetX - xOffset * 2));

            Rect parentBounds = tasklistcontainer.Current.BoundingRectangle;
            Double deltax = Math.Abs(targetX - tasklistBounds.X);

            // Previous bounds check
            if (deltax <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call (DeltaX = " + deltax + ")");
                return; 
            }

            // Right bounds check
            double safeRight = barWidth - parentBounds.Right;
            if ((targetX + width) > (barWidth - safeRight))
            {
                // Shift off center when the bar is too big
                Debug.WriteLine("Shifting off center, too big and hitting right boundary (" + (targetX + width) + " > " + (barWidth - safeRight) + ")");
                double extra = (targetX + tasklistBounds.Width) - (barWidth - safeRight);
                targetX -= extra;
            }

            // Left bounds check
            if (targetX <= (parentBounds.X + ((tasklistcontainer.Current.ClassName == Shell_SecondaryTrayWnd) ? 96: 0)))
            {
                // Prevent X position ending up beyond the normal left aligned position
                Debug.WriteLine("Target is more left than left aligned default, left aligning (" + targetX + " <= " + parentBounds.X + ")");
                Reset(trayWnd);
            }

            int oldWidth = (int)tasklistBounds.Width;
            int oldHeight = (int)tasklistBounds.Height;

            double dx = targetX - tasklist.Current.BoundingRectangle.X;
            targets[trayWnd.Current.ClassName] = targetX;
            if (running.ContainsKey(trayWnd.Current.ClassName) && running[trayWnd.Current.ClassName]) {
                Debug.WriteLine("Updating target while animation is running");
                return;
            }
            running[trayWnd.Current.ClassName] = true;
            int fps = 60;

            // Animation Loop
            while (Math.Abs(dx) >= 2)
         
            {
                double v = safeVelocity(dx / (fps/6));
                double x = tasklist.Current.BoundingRectangle.X + v;
                if ((tasklist.Current.BoundingRectangle.X < targets[trayWnd.Current.ClassName] && x < tasklist.Current.BoundingRectangle.X) || (tasklist.Current.BoundingRectangle.X > targets[trayWnd.Current.ClassName] && x > tasklist.Current.BoundingRectangle.X)) {
                    break;
                }               
                MoveWindow(tasklistPtr, relativeX(x, tasklistcontainer), 0, oldWidth, oldHeight, true);
                Thread.Sleep(1000/fps);
                dx = targets[trayWnd.Current.ClassName] - tasklist.Current.BoundingRectangle.X;
                // Debug.WriteLine(x + "(" + tasklist.Current.BoundingRectangle.X + "), " + v + ", " + dx + " T: " + targets[trayWnd.Current.ClassName]);
            }

            MoveWindow(tasklistPtr, relativeX(targets[trayWnd.Current.ClassName], tasklistcontainer), 0, oldWidth, oldHeight, true);

            running[trayWnd.Current.ClassName] = false;

            Debug.Write("Final X Position: ");
            Debug.WriteLine(tasklist.Current.BoundingRectangle.X);
            Debug.Write((tasklist.Current.BoundingRectangle.X == targets[trayWnd.Current.ClassName]) ? "Move hit target": "Move missed target");
            Debug.WriteLine(" (diff: " + Math.Abs(tasklist.Current.BoundingRectangle.X - targets[trayWnd.Current.ClassName]) + ")");
        }

        private double safeVelocity(double v)
        {
            double minV = 2;
            if (v > 0 && v < minV)
            {
                return minV;
            }
            else if (v < 0 && v > -minV)
            {
                return -minV;
            }
            return Math.Round(v);
        }

        private int relativeX(double x, AutomationElement parent)
        {
            Rect barBounds = parent.Current.BoundingRectangle;
            Double adjustment = (parent.Current.ClassName == Shell_SecondaryTrayWnd) ? 96 : 0; // SecondaryTrayWnd seems to have padding of 96
            Double newPos = x - barBounds.X - adjustment;
            
            if (newPos < 0)
            {
                Debug.WriteLine("Relative position < 0, adjusting to 0 (Previous: " + newPos + ")");
                newPos = 0;
            }

            return (int)newPos;
        }
    }
}
