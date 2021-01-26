using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Automation;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Collections.Generic;
using System.Reflection;

namespace CenterTaskbar
{
    static class Program
    {
        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            // Only allow one instance of this application to run at a time using GUID
            string assyGuid = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value.ToUpper();
            using (Mutex mutex = new Mutex(true, assyGuid, out bool firstInstance))
            {
                if (!firstInstance)
                {
                    MessageBox.Show("Another instance is already running.", "CenterTaskbar", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayApplication(args));
            }
        }
    }

    public class TrayApplication : ApplicationContext
    {
        private const String appName = "CenterTaskbar";
        private const String runRegkey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOZORDER = 0x0004;
        //private const int SWP_SHOWWINDOW = 0x0040;
        private const int SWP_ASYNCWINDOWPOS = 0x4000;
        private const String MSTaskListWClass = "MSTaskListWClass";
        //private const String ReBarWindow32 = "ReBarWindow32";
        private const String Shell_TrayWnd = "Shell_TrayWnd";
        private const String Shell_SecondaryTrayWnd = "Shell_SecondaryTrayWnd";
        private static readonly String ExecutablePath = "\"" + Application.ExecutablePath + "\"";

        private static bool disposed = false;

        private readonly NotifyIcon trayIcon;
        private static readonly AutomationElement desktop = AutomationElement.RootElement;
        private static AutomationEventHandler UIAeventHandler;
        private static AutomationPropertyChangedEventHandler propChangeHandler;
        private readonly Dictionary<AutomationElement, double> lasts = new Dictionary<AutomationElement, double>();
        private readonly Dictionary<AutomationElement, AutomationElement> children = new Dictionary<AutomationElement, AutomationElement>();
        private readonly List<AutomationElement> bars = new List<AutomationElement>();

        private readonly int activeFramerate = DisplaySettings.CurrentRefreshRate();
        // private Thread positionThread;
        private Dictionary<AutomationElement, Thread> positionThreads = new Dictionary<AutomationElement, Thread>();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        public TrayApplication(string[] args)
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

            MenuItem header = new MenuItem("CenterTaskbar (" + activeFramerate + ")", Exit)
            {
                Enabled = false
            };

            MenuItem startup = new MenuItem("Start with Windows", ToggleStartup)
            {
                Checked = IsApplicationInStatup()
            };

            // Setup Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.TrayIcon,
                ContextMenu = new ContextMenu(new MenuItem[] {
                    header,
                    new MenuItem("Scan for screens", Restart),
                    startup,
                    new MenuItem("E&xit", Exit)
                }),
                Visible = true
            };

            Start();

            SystemEvents.DisplaySettingsChanging += SystemEvents_DisplaySettingsChanged;
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
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(runRegkey, true))
            {
                if (key == null) return false;

                object value = key.GetValue(appName);
                if (value is String)
                {
                    ;
                    return ((value as String).StartsWith(ExecutablePath));
                }

                return false;
            }
        }

        public void AddApplicationToStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(runRegkey, true))
            {
                key.SetValue(appName, ExecutablePath);
            }
        }

        public void RemoveApplicationFromStartup()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(runRegkey, true))
            {
                key.DeleteValue(appName, false);
            }
        }

        private void Exit(object sender, EventArgs e)
        {
            SystemEvents.DisplaySettingsChanging -= SystemEvents_DisplaySettingsChanged;
            Application.ExitThread();
        }

        private void CancelPositionThread()
        {
            foreach (Thread thread in positionThreads.Values)
            {
                thread.Abort();
            }
        }

        void Restart(object sender, EventArgs e)
        {
            CancelPositionThread();
            Start();
        }

        private void ResetAll()
        {
            CancelPositionThread();
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

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;
            
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

                    propChangeHandler = new AutomationPropertyChangedEventHandler(OnUIAutomationEvent);
                    Automation.AddAutomationPropertyChangedEventHandler(tasklist, TreeScope.Element, propChangeHandler, AutomationElement.BoundingRectangleProperty);

                    bars.Add(trayWnd);
                    children.Add(trayWnd, tasklist);
                    positionThreads[trayWnd] = new Thread(new ParameterizedThreadStart(LoopForPosition));
                    positionThreads[trayWnd].Start(trayWnd);
                }
            }

            UIAeventHandler = new AutomationEventHandler(OnUIAutomationEvent);
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, TreeScope.Subtree, UIAeventHandler);
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, TreeScope.Subtree, UIAeventHandler);
        }

        private void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
        {
            Restart(sender, e);
        }

        private void OnUIAutomationEvent(object src, AutomationEventArgs e)
        {
            Debug.Print("Event occured: {0}", e.EventId.ProgrammaticName);
            foreach (AutomationElement trayWnd in bars)
            {
                if (!positionThreads[trayWnd].IsAlive)
                {
                    Debug.WriteLine("Starting new thead");
                    positionThreads[trayWnd] = new Thread(new ParameterizedThreadStart(LoopForPosition));
                    positionThreads[trayWnd].Start(trayWnd);
                } else
                {
                    Debug.WriteLine("Thread already exists");
                }
            }

        }

        private void LoopForPosition(object trayWndObj)
        {
            AutomationElement trayWnd = (AutomationElement)trayWndObj;
            int numberOfLoops = activeFramerate / 10;
            int keepGoing = 0;
            while (keepGoing < numberOfLoops)
            {
                if (!PositionLoop(trayWnd))
                {
                    keepGoing += 1;
                }
                System.Threading.Tasks.Task.Delay(1000 / activeFramerate).Wait();             
            }

            Debug.WriteLine("LoopForPosition Thread ended.");
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
            bool horizontal = trayBounds.Width > trayBounds.Height;

            // Use the left/top bounds because there is an empty element as the last child with a nonzero width
            double lastChildPos = horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top; 
            Debug.WriteLine("Last child position: " + lastChildPos);

            if (lasts.ContainsKey(trayWnd) && lastChildPos == lasts[trayWnd])
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

                int iconSizeSetting = (int)Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSmallIcons", 0);
                int iconSizeHorizontal = (iconSizeSetting == 0) ? 40 : 30;
                int iconSizeVertical = (iconSizeSetting == 0) ? 47 : 31;

                double scale = horizontal
                    ? first.Current.BoundingRectangle.Height / iconSizeHorizontal
                    : first.Current.BoundingRectangle.Height / iconSizeVertical;
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
                int rightBounds = SideBoundary(false, horizontal, tasklist, scale, trayBounds);
                if ((targetPos + size) > rightBounds)
                {
                    // Shift off center when the bar is too big
                    double extra = (targetPos + size) - rightBounds;
                    Debug.WriteLine("Shifting off center, too big and hitting right/bottom boundary (" + (targetPos + size) + " > " + rightBounds + ") // " + extra);
                    targetPos -= extra;
                }

                // Left bounds check
                int leftBounds = SideBoundary(true, horizontal, tasklist, scale, trayBounds);
                if (targetPos <= leftBounds)
                {
                    // Prevent X position ending up beyond the normal left aligned position
                    Debug.WriteLine("Target is more left than left/top aligned default, left/top aligning (" + targetPos + " <= " + leftBounds + ")");
                    Reset(trayWnd);
                    return true;
                }

                IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

                if (horizontal)
                {
                    SetWindowPos(tasklistPtr, IntPtr.Zero, (RelativePos(targetPos, horizontal, tasklist, scale, trayBounds)), 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
                    Debug.Write("Final X Position: ");
                    Debug.WriteLine(((first.Current.BoundingRectangle.Left - trayBounds.Left) / scale) + trayBounds.Left);
                    Debug.Write(Math.Round(((first.Current.BoundingRectangle.Left - trayBounds.Left) / scale) + trayBounds.Left) == Math.Round(targetPos) ? "Move hit target" : "Move missed target");
                    Debug.WriteLine(" (diff: " + Math.Abs((((first.Current.BoundingRectangle.Left - trayBounds.Left) / scale) + trayBounds.Left) - targetPos) + ")");
                } else
                {
                    SetWindowPos(tasklistPtr, IntPtr.Zero, 0, (RelativePos(targetPos, horizontal, tasklist, scale, trayBounds)), 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
                    Debug.Write("Final Y Position: ");
                    Debug.WriteLine(((first.Current.BoundingRectangle.Top - trayBounds.Top) / scale) + trayBounds.Top);
                    Debug.Write(Math.Round(((first.Current.BoundingRectangle.Top - trayBounds.Top) / scale) + trayBounds.Top) == Math.Round(targetPos) ? "Move hit target" : "Move missed target");
                    Debug.WriteLine(" (diff: " + Math.Abs((((first.Current.BoundingRectangle.Top - trayBounds.Top) / scale) + trayBounds.Top) - targetPos) + ")");
                }

                lasts[trayWnd] = horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top;

                return true;
            }
        }

        private int RelativePos(double x, bool horizontal, AutomationElement element, double scale, System.Windows.Rect trayBounds)
        {
            int adjustment = SideBoundary(true, horizontal, element, scale, trayBounds);

            double newPos = x - adjustment;

            if (newPos < 0)
            {
                Debug.WriteLine("Relative position < 0, adjusting to 0 (Previous: " + newPos + ")");
                newPos = 0;
            }

            return (int)newPos;
        }

        private int SideBoundary(bool left, bool horizontal, AutomationElement element, double scale, System.Windows.Rect trayBounds)
        {
            double adjustment = 0;
            Debug.WriteLine("Boundary calc for " + element.Current.ClassName);
            var prevSibling = TreeWalker.RawViewWalker.GetPreviousSibling(element);
            var nextSibling = TreeWalker.RawViewWalker.GetNextSibling(element);
            var first = TreeWalker.RawViewWalker.GetFirstChild(element);
            var parent = TreeWalker.RawViewWalker.GetParent(element);

            var padding = horizontal ? (trayBounds.Left - element.Current.BoundingRectangle.Left) - ((trayBounds.Left - first.Current.BoundingRectangle.Left) / scale) : (trayBounds.Top - element.Current.BoundingRectangle.Top) - ((trayBounds.Top - first.Current.BoundingRectangle.Top) / scale);

            Debug.Write(horizontal ? "Horizontal Padding: " : "Vertical Padding: ");
            Debug.WriteLine(Math.Round(padding));
            if (padding < 0)
            {
                Debug.WriteLine("Padding should not be less than 0, setting to 0");
                padding = 0;
            }

            if ((left && prevSibling != null && !prevSibling.Current.BoundingRectangle.IsEmpty))
            {
                Debug.WriteLine("Left sibling calc " + prevSibling.Current.ClassName);
                adjustment = horizontal ? prevSibling.Current.BoundingRectangle.Right : prevSibling.Current.BoundingRectangle.Bottom;
            } else if (!left && nextSibling != null && !nextSibling.Current.BoundingRectangle.IsEmpty)
            {
                adjustment = horizontal ? nextSibling.Current.BoundingRectangle.Left : nextSibling.Current.BoundingRectangle.Top;
            }
            else if (parent != null)
            {
                Debug.WriteLine("Right sibling calc " + nextSibling.Current.ClassName);
                if (horizontal)
                {
                    adjustment = left ? parent.Current.BoundingRectangle.Left + padding : parent.Current.BoundingRectangle.Right;
                } else
                {
                    adjustment = left ? parent.Current.BoundingRectangle.Top + padding : parent.Current.BoundingRectangle.Bottom;
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

        // Protected implementation of Dispose pattern.
        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // Stop listening for new events
                if (UIAeventHandler != null)
                {
                    //Automation.RemoveAutomationPropertyChangedEventHandler(tasklist, propChangeHandler);  //TODO: Remove these from each taskbar
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, desktop, UIAeventHandler);
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowClosedEvent, desktop, UIAeventHandler);
                }

                // Put icons back
                ResetAll();

                // Hide tray icon, otherwise it will remain shown until user mouses over it
                trayIcon.Visible = false;                
                trayIcon.Dispose();
            }

            disposed = true;
        }
    }
}
