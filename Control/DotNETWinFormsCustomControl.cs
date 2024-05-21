using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;

// IMPORTANT: must match the typelib's GUID
[assembly: Guid("46F3FEB2-121D-4830-AA22-0CDA9EA90DC3")]

// The control will appear under the name "<Namespace>.<ClassName>".
// In our case it will be "DotNETWinFormsCustomControlLib.DotNETWinFormsCustomControl"
namespace DotNETWinFormsCustomControlLib
{
    [ComVisible(true)]
    [Guid("55E09A56-35A5-48AF-B1EC-B63AF0BAEEA5")]
    public partial class DotNETWinFormsCustomControl : UserControl
    {
        public DotNETWinFormsCustomControl()
        {
            InitializeComponent();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Eureca!");
        }

        private static readonly string TlbPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!, $"Control.comhost.dll");

        [ComRegisterFunction()]
        public static void RegisterClass(string key)
        {
            // Registering typelib
            TypeLibRegistration.Register(TlbPath);

            using var clsidKey = Registry.LocalMachine.OpenSubKey(key.Replace(@"HKEY_LOCAL_MACHINE\", ""), true)!;

            // Marking as "Control", so that VBA can see it
            using (clsidKey.CreateSubKey("Control")) { }
        }

        [ComUnregisterFunction()]
        public static void UnregisterClass(string key)
        {
            using var clsidKey = Registry.LocalMachine.OpenSubKey(key.Replace(@"HKEY_LOCAL_MACHINE\", ""), true);
            if (clsidKey != null)
            {
                clsidKey.DeleteSubKey("Control", false);
            }

            // Unregistering typelib
            TypeLibRegistration.Unregister(TlbPath);
        }
    }
}
