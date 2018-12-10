using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Runtime.InteropServices;

namespace CenterTaskbar
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
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
        static AutomationElement taskbar = desktop.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "ReBarWindow32"));

        Double padding = 2; // 2px space between ReBarWindow32 and MSTaskListWClass apparently
        int screenWidth;
        IntPtr tasklistPtr;
        Double safeRight;

        [DllImport("user32.dll")]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        public CustomApplicationContext()
        {
            // Setup Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.Icon1,
                ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };

            tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;
            screenWidth = Screen.PrimaryScreen.Bounds.Width;
            safeRight = screenWidth - (taskbar.Current.BoundingRectangle.Width + taskbar.Current.BoundingRectangle.X);

            Reposition();

            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Subtree, OnPropertyChange);
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, AutomationElement.RootElement, TreeScope.Subtree, OnPropertyChange);
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;
            Reset();
            System.Windows.Forms.Application.Exit();
        }

        private void OnPropertyChange(object sender, AutomationEventArgs e)
        {
            // Windows changed, update position
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

        private void Reset()
        {
            if (!SafetyCheck())
            {
                return;
            }

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect barBounds = taskbar.Current.BoundingRectangle;
            Double deltax = (listBounds.X - padding) - barBounds.X;

            if (Math.Abs(deltax) <= padding)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                return;
            }

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;
            Rect bounds = tasklist.Current.BoundingRectangle;
            int newWidth = (int)bounds.Width;
            int newHeight = (int)bounds.Height;
            MoveWindow(tasklistPtr, (int)padding, 0, newWidth, newHeight, true);
        }

        private void Reposition()
        {
            if (!SafetyCheck())
            {
                return;
            }

            AutomationElementCollection children = tasklist.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
            if (children == null)
            {
                return;
            }

            Double width = 0;
            foreach (AutomationElement element in children)
            {
                width += element.Current.BoundingRectangle.Width;
            }

            double targetX = (screenWidth / 2) - (width / 2);

            Rect listBounds = tasklist.Current.BoundingRectangle;
            Rect barBounds = taskbar.Current.BoundingRectangle;
            Double deltax = targetX - (listBounds.X - padding);

            Console.WriteLine(deltax);

            if (Math.Abs(deltax) <= padding)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                return; 
            }

            if ((targetX + width) > (screenWidth - safeRight))
            {
                // Shift off center when the bar is too big
                double extra = (targetX + listBounds.Width) - (screenWidth - safeRight);
                targetX -= extra;
            }

            if (targetX <= (barBounds.X + padding))
            {
                // Prevent X position ending up beyond the normal left aligned position
                Reset();
            }

            Double newPos = targetX - barBounds.X; // MoveWindow uses relative positioning
            int oldWidth = (int)listBounds.Width;
            int oldHeight = (int)listBounds.Height;

            if (newPos < padding)
            {
                newPos = padding;
            }
 
            MoveWindow(tasklistPtr, (int)newPos, 0, oldWidth, oldHeight, true);
        }
    }
}
