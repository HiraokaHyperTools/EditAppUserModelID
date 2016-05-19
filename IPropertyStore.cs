using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace EditAppUserModelID {
    [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore {
        void GetCount([Out] out uint cProps);

        void GetAt([In] uint iProp, ref PropertyKey pkey);

        void GetValue([In] ref PropertyKey key, ref PropVariant pv);

        void SetValue([In] ref PropertyKey key, [In] ref PropVariant pv);

        void Commit();
    }

    [StructLayout(LayoutKind.Sequential, Size = 32)]
    public struct PropertyKey {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Sequential, Size = 32)]
    public struct PropVariant {
        public ushort vt;
        ushort wReserved1;
        ushort wReserved2;
        ushort wReserved3;
        public IntPtr pVal;
    }
}
