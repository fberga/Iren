using System;
using System.Windows.Forms;

namespace Iren.PSO.Base
{
    public class Win32Window : IWin32Window
    {
        public Win32Window(IntPtr handle) { Handle = handle; }
        public IntPtr Handle { get; private set; }
    }
}
