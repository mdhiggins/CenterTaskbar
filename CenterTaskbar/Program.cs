using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using CenterTaskbar.Properties;
using Microsoft.Win32;

namespace CenterTaskbar
{
    internal static class Program
    {
        /// <summary>
        ///     The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main(string[] args)
        {
            // Only allow one instance of this application to run at a time using GUID
            var assyGuid = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value.ToUpper();
            using (new Mutex(true, assyGuid, out var firstInstance))
            {
                if (!firstInstance)
                {
                    MessageBox.Show("Another instance is already running.", "CenterTaskbar", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
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
        private const string AppName = "CenterTaskbar";
        private const string RunRegkey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
        private const int OneSecond = 1000;
        private const int SWP_NOSIZE = 0x0001;
        private const int SWP_NOZORDER = 0x0004;
        //private const int SWP_SHOWWINDOW = 0x0040;
        private const int SWP_ASYNCWINDOWPOS = 0x4000;
        private const string MSTaskListWClass = "MSTaskListWClass";
        //private const String ReBarWindow32 = "ReBarWindow32";
        private const string ShellTrayWnd = "Shell_TrayWnd";
        private const string ShellSecondaryTrayWnd = "Shell_SecondaryTrayWnd";

        private static readonly string ExecutablePath = "\"" + Application.ExecutablePath + "\"";
        private static bool _disposed;
        private CancellationTokenSource _loopCancellationTokenSource = new CancellationTokenSource();

        private static readonly AutomationElement Desktop = AutomationElement.RootElement;
        private static AutomationEventHandler _uiaEventHandler;
        private static AutomationPropertyChangedEventHandler _propChangeHandler;

        private readonly int _activeFramerate = DisplaySettings.CurrentRefreshRate();
        private readonly List<AutomationElement> _bars = new List<AutomationElement>();

        private readonly Dictionary<AutomationElement, AutomationElement> _children =
            new Dictionary<AutomationElement, AutomationElement>();

        private readonly Dictionary<AutomationElement, double> _lasts = new Dictionary<AutomationElement, double>();

        private readonly NotifyIcon _trayIcon;

        // private Thread positionThread;
        private readonly Dictionary<AutomationElement, Task> _positionThreads =
            new Dictionary<AutomationElement, Task>();

        public TrayApplication(IReadOnlyList<string> args)
        {
            if (args.Count > 0)
                try
                {
                    _activeFramerate = int.Parse(args[0]);
                    Debug.WriteLine("Active refresh rate: " + _activeFramerate);
                }
                catch (FormatException e)
                {
                    Debug.WriteLine(e.Message);
                }

            var header = new MenuItem("CenterTaskbar (" + _activeFramerate + ")", Exit)
            {
                Enabled = false
            };

            var startup = new MenuItem("Start with Windows", ToggleStartup)
            {
                Checked = IsApplicationInStartup()
            };

            // Setup Tray Icon
            _trayIcon = new NotifyIcon
            {
                Icon = Resources.TrayIcon,
                ContextMenu = new ContextMenu(new[]
                {
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

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, int uFlags);

        public void ToggleStartup(object sender, EventArgs e)
        {
            if (IsApplicationInStartup())
            {
                RemoveApplicationFromStartup();
                ((MenuItem) sender).Checked = false;
            }
            else
            {
                AddApplicationToStartup();
                ((MenuItem) sender).Checked = true;
            }
        }

        public bool IsApplicationInStartup()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RunRegkey, true))
            {
                var value = key?.GetValue(AppName);
                return value is string startValue && startValue.StartsWith(ExecutablePath);
            }
        }

        public void AddApplicationToStartup()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RunRegkey, true))
            {
                key?.SetValue(AppName, ExecutablePath);
            }
        }

        public void RemoveApplicationFromStartup()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(RunRegkey, true))
            {
                key?.DeleteValue(AppName, false);
            }
        }

        private void Exit(object sender, EventArgs e)
        {
            SystemEvents.DisplaySettingsChanging -= SystemEvents_DisplaySettingsChanged;
            Application.ExitThread();
        }

        private void CancelPositionThread()
        {
            try
            {
                _loopCancellationTokenSource.Cancel();
                Parallel.ForEach(_positionThreads.Values.ToList(), theTask =>
                {
                    try
                    {
                        // Give the thread time to exit gracefully.
                        if (theTask.Wait(OneSecond * 3)) return;
                    }
                    catch (OperationCanceledException e)
                    {
                        Console.WriteLine($"{nameof(OperationCanceledException)} thrown with message: {e.Message}");
                    }
                    finally
                    {
                        theTask.Dispose();
                    }
                });
            }
            catch (OperationCanceledException e)
            {
                Console.WriteLine($"{nameof(OperationCanceledException)} thrown with message: {e.Message}");
            }
            finally
            {
                _loopCancellationTokenSource = new CancellationTokenSource();
            }
        }

        private void Restart(object sender, EventArgs e)
        {
            CancelPositionThread();
            Start();
        }

        private void ResetAll()
        {
            CancelPositionThread();
            Parallel.ForEach(_bars.ToList(), Reset);
        }

        private static void Reset(AutomationElement trayWnd)
        {
            Debug.WriteLine("Begin Reset Calculation");

            var taskList = trayWnd.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass));
            if (taskList == null)
            {
                Debug.WriteLine("Null values found, aborting reset");
                return;
            }

            var taskListContainer = TreeWalker.ControlViewWalker.GetParent(taskList);
            if (taskListContainer == null)
            {
                Debug.WriteLine("Null values found, aborting reset");
                return;
            }

            var taskListPtr = (IntPtr) taskList.Current.NativeWindowHandle;

            SetWindowPos(taskListPtr, IntPtr.Zero, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
        }

        private void Start()
        {
            var condition = new OrCondition(new PropertyCondition(AutomationElement.ClassNameProperty, ShellTrayWnd),
                new PropertyCondition(AutomationElement.ClassNameProperty, ShellSecondaryTrayWnd));
            var cacheRequest = new CacheRequest();
            cacheRequest.Add(AutomationElement.NameProperty);
            cacheRequest.Add(AutomationElement.BoundingRectangleProperty);

            _bars.Clear();
            _children.Clear();
            _lasts.Clear();

            using (cacheRequest.Activate())
            {
                var lists = Desktop.FindAll(TreeScope.Children, condition);
                if (lists == null)
                {
                    Debug.WriteLine("Null values found, aborting");
                    return;
                }

                Debug.WriteLine(lists.Count + " bar(s) detected");
                _lasts.Clear();
                Parallel.ForEach(lists.OfType<AutomationElement>(), trayWnd =>
                {
                    var taskList = trayWnd.FindFirst(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.ClassNameProperty, MSTaskListWClass));
                    if (taskList == null)
                    {
                        Debug.WriteLine("Null values found, aborting");
                    }
                    else
                    {
                        _propChangeHandler = OnUIAutomationEvent;
                        Automation.AddAutomationPropertyChangedEventHandler(taskList, TreeScope.Element, _propChangeHandler,
                            AutomationElement.BoundingRectangleProperty);

                        _bars.Add(trayWnd);
                        _children.Add(trayWnd, taskList);

                        _positionThreads[trayWnd] = Task.Run(() => LoopForPosition(trayWnd), _loopCancellationTokenSource.Token);
                    }
                });
            }

            _uiaEventHandler = OnUIAutomationEvent;
            Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, Desktop, TreeScope.Subtree, _uiaEventHandler);
            Automation.AddAutomationEventHandler(WindowPattern.WindowClosedEvent, Desktop, TreeScope.Subtree, _uiaEventHandler);
        }

        private async void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
        {
            await Task.Delay(3000);
            Restart(sender, e);
        }

        private void OnUIAutomationEvent(object src, AutomationEventArgs e)
        {
            Debug.Print("Event occured: {0}", e.EventId.ProgrammaticName);
            Parallel.ForEach(_bars.ToList(), trayWnd => 
            {
                if (_positionThreads[trayWnd].IsCompleted)
                {
                    Debug.WriteLine("Starting new thead");
                    _positionThreads[trayWnd] = Task.Run(() => LoopForPosition(trayWnd), _loopCancellationTokenSource.Token);
                }
                else
                {
                    Debug.WriteLine("Thread already exists");
                }
            });
        }

        private void LoopForPosition(object trayWndObj)
        {
            var trayWnd = (AutomationElement) trayWndObj;
            var numberOfLoops = _activeFramerate / 10;
            var keepGoing = 0;
            while (keepGoing < numberOfLoops)
            {
                if (!PositionLoop(trayWnd)) keepGoing += 1;
                if (_loopCancellationTokenSource.IsCancellationRequested) break;
                Task.Delay(OneSecond / _activeFramerate).Wait();
            }

            Debug.WriteLine("LoopForPosition Thread ended.");
        }

        private bool PositionLoop(AutomationElement trayWnd)
        {
            Debug.WriteLine("Begin Reposition Calculation");

            var taskList = _children[trayWnd];
            var last = TreeWalker.ControlViewWalker.GetLastChild(taskList);
            if (last == null)
            {
                Debug.WriteLine("Null values found for items, aborting");
                return true;
            }

            var trayBounds = trayWnd.Cached.BoundingRectangle;
            var horizontal = trayBounds.Width > trayBounds.Height;

            // Use the left/top bounds because there is an empty element as the last child with a nonzero width
            var lastChildPos = horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top;
            Debug.WriteLine("Last child position: " + lastChildPos);

            if (_lasts.ContainsKey(trayWnd) && lastChildPos == _lasts[trayWnd])
            {
                Debug.WriteLine("Size/location unchanged, sleeping");
                return false;
            }

            Debug.WriteLine("Size/location changed, recalculating center");
            _lasts[trayWnd] = lastChildPos;

            var first = TreeWalker.ControlViewWalker.GetFirstChild(taskList);
            if (first == null)
            {
                Debug.WriteLine("Null values found for first child item, aborting");
                return true;
            }

            var iconSizeSetting = (int)Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "TaskbarSmallIcons", 0);
            var iconSizeHorizontal = (iconSizeSetting == 0) ? 40 : 30;
            var iconSizeVertical = (iconSizeSetting == 0) ? 47 : 31;

            var scale = horizontal
                ? first.Current.BoundingRectangle.Height / iconSizeHorizontal
                : first.Current.BoundingRectangle.Height / iconSizeVertical;
            Debug.WriteLine("UI Scale: " + scale);
            var size = (lastChildPos - (horizontal 
                            ? first.Current.BoundingRectangle.Left 
                            : first.Current.BoundingRectangle.Top)
                        ) / scale;
            if (size < 0)
            {
                Debug.WriteLine("Size calculation failed");
                return true;
            }

            var taskListContainer = TreeWalker.ControlViewWalker.GetParent(taskList);
            if (taskListContainer == null)
            {
                Debug.WriteLine("Null values found for parent, aborting");
                return true;
            }

            var taskListBounds = taskList.Current.BoundingRectangle;

            var barSize = horizontal ? trayWnd.Cached.BoundingRectangle.Width : trayWnd.Cached.BoundingRectangle.Height;
            var targetPos = Math.Round((barSize - size) / 2) + (horizontal ? trayBounds.X : trayBounds.Y);

            Debug.Write("Bar size: ");
            Debug.WriteLine(barSize);
            Debug.Write("Total icon size: ");
            Debug.WriteLine(size);
            Debug.Write("Target abs " + (horizontal ? "X" : "Y") + " position: ");
            Debug.WriteLine(targetPos);

            var delta = Math.Abs(targetPos - (horizontal ? taskListBounds.X : taskListBounds.Y));
            // Previous bounds check
            if (delta <= 1)
            {
                // Already positioned within margin of error, avoid the unneeded MoveWindow call
                Debug.WriteLine("Already positioned, ending to avoid the unneeded MoveWindow call (Delta: " + delta + ")");
                return false;
            }

            // Right bounds check
            int rightBounds;
            int leftBounds;
            try
            {
                rightBounds = SideBoundary(false, horizontal, taskList, scale, trayBounds);
                leftBounds = SideBoundary(true, horizontal, taskList, scale, trayBounds);
            }
            catch (NullReferenceException)
            {
                Reset(trayWnd);
                return true;
            }
            
            if (targetPos + size > rightBounds)
            {
                // Shift off center when the bar is too big
                var extra = targetPos + size - rightBounds;
                Debug.WriteLine("Shifting off center, too big and hitting right/bottom boundary (" + (targetPos + size) + " > " + rightBounds + ") // " + extra);
                targetPos -= extra;
            }

            // Left bounds check
            if (targetPos <= leftBounds)
            {
                // Prevent X position ending up beyond the normal left aligned position
                Debug.WriteLine("Target is more left than left/top aligned default, left/top aligning (" + targetPos + " <= " + leftBounds + ")");
                Reset(trayWnd);
                return true;
            }

            var taskListPtr = (IntPtr) taskList.Current.NativeWindowHandle;

            if (horizontal)
            {
                SetWindowPos(taskListPtr, IntPtr.Zero, RelativePos(targetPos, horizontal, taskList, scale, trayBounds), 0, 0, 0,
                    SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
                Debug.Write("Final X Position: ");
                Debug.WriteLine(((first.Current.BoundingRectangle.Left - trayBounds.Left) / scale) + trayBounds.Left);
                Debug.Write(Math.Round(((first.Current.BoundingRectangle.Left - trayBounds.Left) / scale) + trayBounds.Left) == Math.Round(targetPos) ? "Move hit target" : "Move missed target");
                Debug.WriteLine(" (diff: " + Math.Abs((((first.Current.BoundingRectangle.Left - trayBounds.Left) / scale) + trayBounds.Left) - targetPos) + ")");
            }
            else
            {
                SetWindowPos(taskListPtr, IntPtr.Zero, 0, RelativePos(targetPos, horizontal, taskList, scale, trayBounds), 0, 0,
                    SWP_NOZORDER | SWP_NOSIZE | SWP_ASYNCWINDOWPOS);
                Debug.Write("Final Y Position: ");
                Debug.WriteLine(((first.Current.BoundingRectangle.Top - trayBounds.Top) / scale) + trayBounds.Top);
                Debug.Write(Math.Round(((first.Current.BoundingRectangle.Top - trayBounds.Top) / scale) + trayBounds.Top) == Math.Round(targetPos) ? "Move hit target" : "Move missed target");
                Debug.WriteLine(" (diff: " + Math.Abs((((first.Current.BoundingRectangle.Top - trayBounds.Top) / scale) + trayBounds.Top) - targetPos) + ")");
            }

            _lasts[trayWnd] = horizontal ? last.Current.BoundingRectangle.Left : last.Current.BoundingRectangle.Top;

            return true;
        }

        private static int RelativePos(double x, bool horizontal, AutomationElement element, double scale, System.Windows.Rect trayBounds)
        {
            var adjustment = SideBoundary(true, horizontal, element, scale, trayBounds);

            var newPos = x - adjustment;

            if (newPos < 0)
            {
                Debug.WriteLine("Relative position < 0, adjusting to 0 (Previous: " + newPos + ")");
                newPos = 0;
            }

            return (int) newPos;
        }

        private static int SideBoundary(bool left, bool horizontal, AutomationElement element, double scale, System.Windows.Rect trayBounds)
        {
            double adjustment = 0;
            Debug.WriteLine("Boundary calc for " + element.Current.ClassName);
            var prevSibling = TreeWalker.RawViewWalker.GetPreviousSibling(element);
            var nextSibling = TreeWalker.RawViewWalker.GetNextSibling(element);
            var first = TreeWalker.RawViewWalker.GetFirstChild(element);
            var parent = TreeWalker.RawViewWalker.GetParent(element);

            var padding = horizontal? (trayBounds.Left - element.Current.BoundingRectangle.Left) - ((trayBounds.Left - first.Current.BoundingRectangle.Left) / scale): (trayBounds.Top - element.Current.BoundingRectangle.Top) - ((trayBounds.Top - first.Current.BoundingRectangle.Top) / scale);

            Debug.Write(horizontal ? "Horizontal Padding: ": "Vertical Padding: ");
            Debug.WriteLine(Math.Round(padding));
            if (padding < 0)
            {
                Debug.WriteLine("Padding should not be less than 0, setting to 0");
                padding = 0;
            }

            if (left && prevSibling != null && !prevSibling.Current.BoundingRectangle.IsEmpty)
            {
                Debug.WriteLine("Left sibling calc " + prevSibling.Current.ClassName);
                adjustment = horizontal
                    ? prevSibling.Current.BoundingRectangle.Right
                    : prevSibling.Current.BoundingRectangle.Bottom;
            }
            else if (!left && nextSibling != null && !nextSibling.Current.BoundingRectangle.IsEmpty)
            {
                Debug.WriteLine("Right sibling calc " + nextSibling.Current.ClassName);
                adjustment = horizontal
                    ? nextSibling.Current.BoundingRectangle.Left
                    : nextSibling.Current.BoundingRectangle.Top;
            }
            else if (parent != null)
            {
                Debug.WriteLine("Parent calc " + parent.Current.ClassName);
                if (horizontal)
                    adjustment = left ? parent.Current.BoundingRectangle.Left + padding : parent.Current.BoundingRectangle.Right;
                else
                    adjustment = left ? parent.Current.BoundingRectangle.Top + padding : parent.Current.BoundingRectangle.Bottom;
            }

            if (horizontal)
                Debug.WriteLine((left ? "Left" : "Right") + " side boundary calculated at " + adjustment);
            else
                Debug.WriteLine((left ? "Top" : "Bottom") + " side boundary calculated at " + adjustment);

            return (int) adjustment;
        }

        // Protected implementation of Dispose pattern.
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                // Stop listening for new events
                if (_uiaEventHandler != null)
                {
                    foreach (var taskBar in _children)
                    {
                        Automation.RemoveAutomationPropertyChangedEventHandler(taskBar.Value, _propChangeHandler); 
                    }

                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowOpenedEvent, Desktop, _uiaEventHandler);
                    Automation.RemoveAutomationEventHandler(WindowPattern.WindowClosedEvent, Desktop, _uiaEventHandler);
                }

                // Put icons back
                ResetAll();

                // Hide tray icon, otherwise it will remain shown until user mouses over it
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
            }

            _disposed = true;
        }
    }
}