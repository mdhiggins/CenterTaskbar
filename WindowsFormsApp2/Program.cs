using System;
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
            Application.Run(new CustomApplicationContext(args));
        }
    }

    public class CustomApplicationContext : ApplicationContext
    {
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

        Dictionary<IntPtr, double> lasts = new Dictionary<IntPtr, double>();

        static int restingFramerate = 10;
        int framerate = restingFramerate;
        int activeFramerate = 60;
        Thread positionThread;

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117,
        }

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

            if (args.Length > 1)
            {
                try
                {
                    restingFramerate = Int32.Parse(args[1]);
                    Debug.WriteLine("Resting refresh rate: " + restingFramerate);
                }
                catch (FormatException e)
                {
                    Debug.WriteLine(e.Message);
                }
            }

                MenuItem header = new MenuItem("CenterTaskbar " + activeFramerate + "/" + restingFramerate, Exit);
            header.Enabled = false;

            // Setup Tray Icon
            trayIcon = new NotifyIcon()
            {
                Icon = Properties.Resources.Icon1,
                ContextMenu = new ContextMenu(new MenuItem[] {
                header,
                new MenuItem("Scan for screens", Restart),
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };

            Start();
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
            OrCondition condition = new OrCondition(new PropertyCondition(AutomationElement.ClassNameProperty, Shell_TrayWnd), new PropertyCondition(AutomationElement.ClassNameProperty, Shell_SecondaryTrayWnd));
            AutomationElementCollection lists = desktop.FindAll(TreeScope.Children, condition);
            if (lists == null)
            {
                Debug.WriteLine("Null values found, aborting");
                return;
            }
            Debug.WriteLine(lists.Count + " bar(s) detected");

            if (positionThread != null)
            {
                positionThread.Abort();
            }

            foreach (AutomationElement trayWnd in lists)
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

            Rect trayBounds = trayWnd.Current.BoundingRectangle;
            bool horizontal = (trayBounds.Width > trayBounds.Height);

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;

            double listBounds = horizontal ? tasklist.Current.BoundingRectangle.X : tasklist.Current.BoundingRectangle.Y;
            double barBounds = horizontal ? tasklistcontainer.Current.BoundingRectangle.X: tasklistcontainer.Current.BoundingRectangle.Y;
            double delta = Math.Abs(listBounds - barBounds);

            if (delta <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call. (Delta = " + delta + ")");
                return;
            }

            Rect bounds = tasklist.Current.BoundingRectangle;
            int newWidth = (int)bounds.Width;
            int newHeight = (int)bounds.Height;
            SetWindowPos(tasklistPtr, IntPtr.Zero, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
        }

        private void Start()
        {
            Stopwatch stopwatch = new Stopwatch();
            OrCondition condition = new OrCondition(new PropertyCondition(AutomationElement.ClassNameProperty, Shell_TrayWnd), new PropertyCondition(AutomationElement.ClassNameProperty, Shell_SecondaryTrayWnd));
            AutomationElementCollection lists = desktop.FindAll(TreeScope.Children, condition);
            if (lists == null)
            {
                Debug.WriteLine("Null values found, aborting");
                return;
            }
            Debug.WriteLine(lists.Count + " bar(s) detected");
            positionThread = new Thread(() =>
            {
                while (true)
                {
                    stopwatch.Start();
                    foreach (AutomationElement trayWnd in lists)
                    {
                        PositionLoop(trayWnd);
                    }
                    stopwatch.Stop();
                    long elapsed = stopwatch.ElapsedMilliseconds;
                    if (elapsed < (1000 / framerate))
                    {
                        Thread.Sleep((int)((1000 / framerate) - elapsed));
                    }
                    stopwatch.Reset();
                }
            });
            positionThread.Start();
        }

        private void PositionLoop(AutomationElement trayWnd)
        {
            Debug.WriteLine("Begin Reposition Calculation");
 
            AutomationElement tasklist = trayWnd.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass));
            if (tasklist == null)
            {
                Debug.WriteLine("Null values found, aborting");
                return;
            }

            AutomationElement last = TreeWalker.ControlViewWalker.GetLastChild(tasklist);
            if (last == null)
            {
                Debug.WriteLine("Null values found for items, aborting");
                return;
            }

            Rect trayBounds = trayWnd.Current.BoundingRectangle;
            bool horizontal = (trayBounds.Width > trayBounds.Height);

            IntPtr tasklistPtr = (IntPtr)tasklist.Current.NativeWindowHandle;
            double lastChildPos = (horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top); // Use the left/top bounds because there is an empty element as the last child with a nonzero width
            Debug.WriteLine("Last child position: " + lastChildPos);

            if ((lasts.ContainsKey(tasklistPtr) && lastChildPos == lasts[tasklistPtr]))
            {
                Debug.WriteLine("Size/location unchanged, sleeping");
                framerate = restingFramerate;
                return;
            } else
            {
                Debug.WriteLine("Size/location changed, recalculating center");
                framerate = activeFramerate;
                lasts[tasklistPtr] = lastChildPos;
                AutomationElement first = TreeWalker.ControlViewWalker.GetFirstChild(tasklist);
                if (last == null)
                {
                    Debug.WriteLine("Null values found for last child item, aborting");
                    return;
                }

                double scale = horizontal ? (last.Current.BoundingRectangle.Height / trayBounds.Height) : (last.Current.BoundingRectangle.Width / trayBounds.Width);
                Debug.WriteLine("UI Scale: " + scale);
                double size = (lastChildPos - (horizontal ? first.Current.BoundingRectangle.Left : first.Current.BoundingRectangle.Top)) / scale;
                if (size <  0)
                {
                    Debug.WriteLine("Size calculation failed");
                    return;
                }

                AutomationElement tasklistcontainer = TreeWalker.ControlViewWalker.GetParent(tasklist);
                if (tasklistcontainer == null)
                {
                    Debug.WriteLine("Null values found for parent, aborting");
                    return;
                }

                Rect tasklistBounds = tasklist.Current.BoundingRectangle;

                double barSize = horizontal ? trayWnd.Current.BoundingRectangle.Width : trayWnd.Current.BoundingRectangle.Height;
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
                    return;
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
                }

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
                lasts[tasklistPtr] = (horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top);
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
