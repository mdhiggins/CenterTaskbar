using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

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
            var assemblyGuid = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value.ToUpper();
            using (new Mutex(true, assemblyGuid, out var firstInstance))
            {
                if (!firstInstance)
                {
                    MessageBox.Show("Another instance is already running.", "CenterTaskbar", MessageBoxButtons.OK,
                        MessageBoxIcon.Exclamation);
                    return;
                }

                Application.SetHighDpiMode(HighDpiMode.SystemAware);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayApplication(args));
            }
        }
    }
}