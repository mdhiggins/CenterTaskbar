using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using Microsoft.Win32;
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
            Application.Run(new CustomApplicationContext(args));
        }
    }

    public class CustomApplicationContext : ApplicationContext
    {
        public const String appName = "CenterTaskbar";

        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOZORDER = 0x0004;
        private const int SWP_SHOWWINDOW = 0x0040;
        private const int SWP_ASYNCWINDOWPOS = 0x4000;

        private NotifyIcon trayIcon;
        static AutomationElement desktop = AutomationElement.RootElement;
        static String MSTaskListWClass = "MSTaskListWClass";
        //static String ReBarWindow32 = "ReBarWindow32";
        static String Shell_TrayWnd = "Shell_TrayWnd";
        static String Shell_SecondaryTrayWnd = "Shell_SecondaryTrayWnd";

        Dictionary<AutomationElement, double> lasts = new Dictionary<AutomationElement, double>();
        Dictionary<AutomationElement, AutomationElement> children = new Dictionary<AutomationElement, AutomationElement>();
        List<AutomationElement> bars = new List<AutomationElement>();

        int activeFramerate = 60;
        Thread positionThread;

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        public CustomApplicationContext(string[] args)
        {
            if (args.Length > 0)
            {
                try
                {
                    activeFramerate = Int32.Parse(args[0]);
                    Debug.WriteLine("Active refresh rate: " + activeFramerate);
                }
                catch (FormatException e)
                {
                    Debug.WriteLine(e.Message);
                }
            }

            MenuItem header = new MenuItem("CenterTaskbar (" + activeFramerate + ")", Exit);
            header.Enabled = false;
            MenuItem startup = new MenuItem("Start with Windows", ToggleStartup);
            startup.Checked = IsApplicationInStatup();

            // Setup Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.Icon1,
                ContextMenu = new ContextMenu(new MenuItem[] {
                header,
                new MenuItem("Scan for screens", Restart),
                startup,
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };

            Start();
        }

        public void ToggleStartup(object sender, EventArgs e)
        {
            if (IsApplicationInStatup())
            {
                RemoveApplicationFromStartup();
                (sender as MenuItem).Checked = false;
            } else
            {
                AddApplicationToStartup();
                (sender as MenuItem).Checked = true;
            }
        }

        public bool IsApplicationInStatup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                if (key == null) return false;

                object value = key.GetValue(appName);
                if (value is String) return ((value as String).StartsWith("\"" + Application.ExecutablePath + "\""));

                return false;
            }
        }

        public void AddApplicationToStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.SetValue(appName, "\"" + Application.ExecutablePath + "\" " + activeFramerate);
            }
        }

        public void RemoveApplicationFromStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                key.DeleteValue(appName, false);
            }
        }

        void Exit(object sender, EventArgs e)
        {
            // Hide tray icon, otherwise it will remain shown until user mouses over it
            trayIcon.Visible = false;
            ResetAll();
            System.Windows.Forms.Application.Exit();
        }

        void Restart(object sender, EventArgs e)
        {
            if (positionThread != null)
            {
                positionThread.Abort();
            }
            Start();
        }

        private void ResetAll()
        {
            if (positionThread != null)
            {
                positionThread.Abort();
            }

            foreach (AutomationElement trayWnd in bars)
            {
                Reset(trayWnd);
            }
        }

        private void Reset(AutomationElement trayWnd)
        {
            Debug.WriteLine("Begin Reset Calculation");

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

            Rect trayBounds = trayWnd.Cached.BoundingRectangle;
            bool horizontal = (trayBounds.Width > trayBounds.Height);

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

            double listBounds = horizontal ? tasklist.Current.BoundingRectangle.X : tasklist.Current.BoundingRectangle.Y;

            Rect bounds = tasklist.Current.BoundingRectangle;
            int newWidth = (int)bounds.Width;
            int newHeight = (int)bounds.Height;
            SetWindowPos(tasklistPtr, IntPtr.Zero, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
        }

        private void Start()
        {
            OrCondition condition = new OrCondition(new PropertyCondition(AutomationElement.ClassNameProperty, Shell_TrayWnd), new PropertyCondition(AutomationElement.ClassNameProperty, Shell_SecondaryTrayWnd));
            CacheRequest cacheRequest = new CacheRequest();
            cacheRequest.Add(AutomationElement.NameProperty);
            cacheRequest.Add(AutomationElement.BoundingRectangleProperty);

            bars.Clear();
            children.Clear();
            lasts.Clear();

            using (cacheRequest.Activate())
            {
                AutomationElementCollection lists = desktop.FindAll(TreeScope.Children, condition);
                if (lists == null)
                {
                    Debug.WriteLine("Null values found, aborting");
                    return;
                }
                Debug.WriteLine(lists.Count + " bar(s) detected");
                lasts.Clear();
                foreach (AutomationElement trayWnd in lists)
                {
                    AutomationElement tasklist = trayWnd.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass));
                    if (tasklist == null)
                    {
                        Debug.WriteLine("Null values found, aborting");
                        continue;
                    }
                    Automation.AddAutomationPropertyChangedEventHandler(tasklist, TreeScope.Element, OnUIAutomationEvent, AutomationElement.BoundingRectangleProperty);

                    bars.Add(trayWnd);
                    children.Add(trayWnd, tasklist);
                }
            }

            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, TreeScope.Subtree, OnUIAutomationEvent);
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, TreeScope.Subtree, OnUIAutomationEvent);
            loop();
        }

        private void OnUIAutomationEvent(object src, AutomationEventArgs e)
        {
            if (!positionThread.IsAlive)
            {
                loop();
            }
        }

        private void loop()
        {
            positionThread = new Thread(() =>
            {
                int keepGoing = 0;
                while (keepGoing < (activeFramerate / 5))
                {
                    foreach (AutomationElement trayWnd in bars)
                    {
                        if (!PositionLoop(trayWnd))
                        {
                            keepGoing += 1;
                        }
                    }
                    Thread.Sleep(1000 / activeFramerate);
                }
                Debug.WriteLine("Thread ended due to inactivity, sleeping");
            });
            positionThread.Start();
        }

        private bool PositionLoop(AutomationElement trayWnd)
        {
            Debug.WriteLine("Begin Reposition Calculation");

            AutomationElement tasklist = children[trayWnd];
            AutomationElement last = TreeWalker.ControlViewWalker.GetLastChild(tasklist);
            if (last == null)
            {
                Debug.WriteLine("Null values found for items, aborting");
                return true;
            }

            Rect trayBounds = trayWnd.Cached.BoundingRectangle;
            bool horizontal = (trayBounds.Width > trayBounds.Height);

            double lastChildPos = (horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top); // Use the left/top bounds because there is an empty element as the last child with a nonzero width
            Debug.WriteLine("Last child position: " + lastChildPos);

            if ((lasts.ContainsKey(trayWnd) && lastChildPos == lasts[trayWnd]))
            {
                Debug.WriteLine("Size/location unchanged, sleeping");
                return false;
            } else
            {
                Debug.WriteLine("Size/location changed, recalculating center");
                lasts[trayWnd] = lastChildPos;

                AutomationElement first = TreeWalker.ControlViewWalker.GetFirstChild(tasklist);
                if (first == null)
                {
                    Debug.WriteLine("Null values found for first child item, aborting");
                    return true;
                }

                double scale = horizontal ? (last.Current.BoundingRectangle.Height / trayBounds.Height) : (last.Current.BoundingRectangle.Width / trayBounds.Width);
                Debug.WriteLine("UI Scale: " + scale);
                double size = (lastChildPos - (horizontal ? first.Current.BoundingRectangle.Left : first.Current.BoundingRectangle.Top)) / scale;
                if (size <  0)
                {
                    Debug.WriteLine("Size calculation failed");
                    return true;
                }

                AutomationElement tasklistcontainer = TreeWalker.ControlViewWalker.GetParent(tasklist);
                if (tasklistcontainer == null)
                {
                    Debug.WriteLine("Null values found for parent, aborting");
                    return true;
                }

                Rect tasklistBounds = tasklist.Current.BoundingRectangle;

                double barSize = horizontal ? trayWnd.Cached.BoundingRectangle.Width : trayWnd.Cached.BoundingRectangle.Height;
                double targetPos = Math.Round((barSize - size) / 2) + (horizontal ? trayBounds.X : trayBounds.Y);

                Debug.Write("Bar size: ");
                Debug.WriteLine(barSize);
                Debug.Write("Total icon size: ");
                Debug.WriteLine(size);
                Debug.Write("Target abs " + (horizontal ? "X":"Y") + " position: ");
                Debug.WriteLine(targetPos);

                double delta = Math.Abs(targetPos - (horizontal ? tasklistBounds.X : tasklistBounds.Y));
                // Previous bounds check
                if (delta <= 1)
                {
                    // Already positioned within margin of error, avoid the unneeded MoveWindow call
                    Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call (Delta: " + delta + ")");
                    return false;
                }

                // Right bounds check
                int rightBounds = sideBoundary(false, horizontal, tasklist);
                if ((targetPos + size) > (rightBounds))
                {
                    // Shift off center when the bar is too big
                    double extra = (targetPos + size) - rightBounds;
                    Debug.WriteLine("Shifting off center, too big and hitting right/bottom boundary (" + (targetPos + size) + " > " + rightBounds + ") // " + extra);
                    targetPos -= extra;
                }

                // Left bounds check
                int leftBounds = sideBoundary(true, horizontal, tasklist);
                if (targetPos <= (leftBounds))
                {
                    // Prevent X position ending up beyond the normal left aligned position
                    Debug.WriteLine("Target is more left than left/top aligned default, left/top aligning (" + targetPos + " <= " + leftBounds + ")");
                    Reset(trayWnd);
                    return true;
                }

                IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

                if (horizontal)
                {
                    SetWindowPos(tasklistPtr, IntPtr.Zero, (relativePos(targetPos, horizontal, tasklist)), 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
                    Debug.Write("Final X Position: ");
                    Debug.WriteLine(tasklist.Current.BoundingRectangle.X);
                    Debug.Write((tasklist.Current.BoundingRectangle.X == targetPos) ? "Move hit target" : "Move missed target");
                    Debug.WriteLine(" (diff: " + Math.Abs(tasklist.Current.BoundingRectangle.X - targetPos) + ")");
                } else
                {
                    SetWindowPos(tasklistPtr, IntPtr.Zero, 0, (relativePos(targetPos, horizontal, tasklist)), 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
                    Debug.Write("Final Y Position: ");
                    Debug.WriteLine(tasklist.Current.BoundingRectangle.Y);
                    Debug.Write((tasklist.Current.BoundingRectangle.Y == targetPos) ? "Move hit target" : "Move missed target");
                    Debug.WriteLine(" (diff: " + Math.Abs(tasklist.Current.BoundingRectangle.Y - targetPos) + ")");
                }
                lasts[trayWnd] = (horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top);

                return true;
            }
        }

        private int relativePos(double x, bool horizontal, AutomationElement element)
        {
            int adjustment = sideBoundary(true, horizontal, element);

            double newPos = x - adjustment;

            if (newPos < 0)
            {
                Debug.WriteLine("Relative position < 0, adjusting to 0 (Previous: " + newPos + ")");
                newPos = 0;
            }

            return (int)newPos;
        }

        private int sideBoundary(bool left, bool horizontal, AutomationElement element)
        {
            double adjustment = 0;
            AutomationElement prevSibling = TreeWalker.ControlViewWalker.GetPreviousSibling(element);
            AutomationElement nextSibling = TreeWalker.ControlViewWalker.GetNextSibling(element);
            AutomationElement parent = TreeWalker.ControlViewWalker.GetParent(element);
            if ((left && prevSibling != null))
            {
                adjustment = (horizontal ? prevSibling.Current.BoundingRectangle.Right : prevSibling.Current.BoundingRectangle.Bottom);
            } else if (!left && nextSibling != null)
            {
                adjustment = (horizontal ? nextSibling.Current.BoundingRectangle.Left : nextSibling.Current.BoundingRectangle.Top);
            }
            else if (parent != null)
            {
                if (horizontal)
                {
                    adjustment = left ? parent.Current.BoundingRectangle.Left : parent.Current.BoundingRectangle.Right;
                } else
                {
                    adjustment = left ? parent.Current.BoundingRectangle.Top : parent.Current.BoundingRectangle.Bottom;
                }
                
            }

            if (horizontal)
            {
                Debug.WriteLine((left ? "Left" : "Right") + " side boundary calulcated at " + adjustment);
            } else
            {
                Debug.WriteLine((left ? "Top" : "Bottom") + " side boundary calulcated at " + adjustment);
            }
            
            return (int)adjustment;
        }
    }
}
