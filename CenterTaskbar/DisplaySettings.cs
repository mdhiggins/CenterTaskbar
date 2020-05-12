using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace CenterTaskbar
{
    internal static class DisplaySettings
    {

        [DllImport("user32.dll")]
        private static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);
        const int ENUM_CURRENT_SETTINGS = -1;
        const int ENUM_REGISTRY_SETTINGS = -2;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private struct DEVMODE
        {
            private const int CCHDEVICENAME = 32;
            private const int CCHFORMNAME = 32;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHDEVICENAME)]
            public string dmDeviceName;
            public short dmSpecVersion;
            public short dmDriverVersion;
            public short dmSize;
            public short dmDriverExtra;
            public int dmFields;
            public int dmPositionX;
            public int dmPositionY;
            public ScreenOrientation dmDisplayOrientation;
            public int dmDisplayFixedOutput;
            public short dmColor;
            public short dmDuplex;
            public short dmYResolution;
            public short dmTTOption;
            public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCHFORMNAME)]
            public string dmFormName;
            public short dmLogPixels;
            public int dmBitsPerPel;
            public int dmPelsWidth;
            public int dmPelsHeight;
            public int dmDisplayFlags;
            public int dmDisplayFrequency;
            public int dmICMMethod;
            public int dmICMIntent;
            public int dmMediaType;
            public int dmDitherType;
            public int dmReserved1;
            public int dmReserved2;
            public int dmPanningWidth;
            public int dmPanningHeight;
        }

        //public static void ListAllDisplayModes()
        //{
        //    DEVMODE vDevMode = new DEVMODE();
        //    int i = 0;
        //    while (EnumDisplaySettings(null, i, ref vDevMode))
        //    {
        //        Console.WriteLine("Width:{0} Height:{1} Color:{2} Frequency:{3}",
        //                                vDevMode.dmPelsWidth,
        //                                vDevMode.dmPelsHeight,
        //                                1 << vDevMode.dmBitsPerPel, vDevMode.dmDisplayFrequency
        //                            );
        //        i++;
        //    }
        //}

        public static int CurrentRefreshRate()
        {
            var vDevMode = new DEVMODE();
            return EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref vDevMode) ? vDevMode.dmDisplayFrequency : 60;
        }
    }
}
