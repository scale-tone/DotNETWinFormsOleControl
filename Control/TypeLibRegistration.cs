using System.Runtime.InteropServices;

using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace DotNETWinFormsCustomControlLib
{
    internal static class TypeLibRegistration
    {
        public static void Register(string tlbPath)
        {
            int hr = OleAut32.LoadTypeLibEx(tlbPath, OleAut32.REGKIND.REGKIND_REGISTER, out ComTypes.ITypeLib typeLib);
            if (hr < 0)
            {
                Marshal.ThrowExceptionForHR(hr);
            }
        }

        public static void Unregister(string tlbPath)
        {
            int hr = OleAut32.LoadTypeLibEx(tlbPath, OleAut32.REGKIND.REGKIND_NONE, out var typeLib);
            if (hr < 0)
            {
                Console.WriteLine($"Unregistering type library failed: 0x{hr:x}");
                return;
            }

            IntPtr attrPtr = IntPtr.Zero;
            try
            {
                typeLib.GetLibAttr(out attrPtr);
                if (attrPtr != IntPtr.Zero)
                {
                    ComTypes.TYPELIBATTR attr = Marshal.PtrToStructure<ComTypes.TYPELIBATTR>(attrPtr);
                    hr = OleAut32.UnRegisterTypeLib(ref attr.guid, attr.wMajorVerNum, attr.wMinorVerNum, attr.lcid, attr.syskind);
                    if (hr < 0)
                    {
                        Console.WriteLine($"Unregistering type library failed: 0x{hr:x}");
                    }
                }
            }
            finally
            {
                if (attrPtr != IntPtr.Zero)
                {
                    typeLib.ReleaseTLibAttr(attrPtr);
                }
            }
        }

        private class OleAut32
        {
            // https://docs.microsoft.com/windows/api/oleauto/ne-oleauto-regkind
            public enum REGKIND
            {
                REGKIND_DEFAULT = 0,
                REGKIND_REGISTER = 1,
                REGKIND_NONE = 2
            }

            // https://docs.microsoft.com/windows/api/oleauto/nf-oleauto-loadtypelibex
            [DllImport(nameof(OleAut32), CharSet = CharSet.Unicode, ExactSpelling = true)]
            public static extern int LoadTypeLibEx(
                [In, MarshalAs(UnmanagedType.LPWStr)] string fileName,
                REGKIND regKind,
                out ComTypes.ITypeLib typeLib);

            // https://docs.microsoft.com/windows/api/oleauto/nf-oleauto-unregistertypelib
            [DllImport(nameof(OleAut32))]
            public static extern int UnRegisterTypeLib(
                ref Guid id,
                short majorVersion,
                short minorVersion,
                int lcid,
                ComTypes.SYSKIND sysKind);
        }
    }
}
