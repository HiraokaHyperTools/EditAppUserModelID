using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

// http://www.vbaccelerator.com/home/NET/Code/Libraries/Shell_Projects/Creating_and_Modifying_Shortcuts/article.asp

namespace EditAppUserModelID {
    [ComImportAttribute()]
    [GuidAttribute("0000010B-0000-0000-C000-000000000046")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPersistFile {
        [PreserveSig]
        int GetClassID(out Guid pClassID);

        [PreserveSig]
        int IsDirty();

        void Load(
            [MarshalAs(UnmanagedType.LPWStr)] string pszFileName,
            uint dwMode
        );

        void Save(
            [MarshalAs(UnmanagedType.LPWStr)] string pszFileName,
            [MarshalAs(UnmanagedType.Bool)] bool fRemember
        );

        void SaveCompleted(
            [MarshalAs(UnmanagedType.LPWStr)] string pszFileName
        );

        void GetCurFile(
            [MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName
        );
    }

    [StructLayoutAttribute(LayoutKind.Sequential, Pack = 4, Size = 0, CharSet = CharSet.Unicode)]
    public struct _WIN32_FIND_DATAW {
        public uint dwFileAttributes;
        public _FILETIME ftCreationTime;
        public _FILETIME ftLastAccessTime;
        public _FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        public string cAlternateFileName;
    }

    [StructLayoutAttribute(LayoutKind.Sequential, Pack = 4, Size = 0)]
    public struct _FILETIME {
        public uint dwLowDateTime;
        public uint dwHighDateTime;
    }

    [GuidAttribute("00021401-0000-0000-C000-000000000046")]
    [ClassInterfaceAttribute(ClassInterfaceType.None)]
    [ComImportAttribute()]
    public class CShellLink { }

    [ComImportAttribute()]
    [GuidAttribute("000214F9-0000-0000-C000-000000000046")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IShellLinkW {
        void GetPath(
            [Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile,
            int cchMaxPath,
            ref _WIN32_FIND_DATAW pfd,
            uint fFlags
        );

        void GetIDList(out IntPtr ppidl);

        void SetIDList(IntPtr pidl);

        void GetDescription(
            [Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile,
            int cchMaxName
        );

        void SetDescription(
            [MarshalAs(UnmanagedType.LPWStr)] string pszName
        );

        void GetWorkingDirectory(
            [Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir,
            int cchMaxPath
        );

        void SetWorkingDirectory(
            [MarshalAs(UnmanagedType.LPWStr)] string pszDir
        );

        void GetArguments(
            [Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs,
            int cchMaxPath
        );

        void SetArguments(
            [MarshalAs(UnmanagedType.LPWStr)] string pszArgs
        );

        void GetHotkey(
            out short pwHotkey
        );

        void SetHotkey(
            short pwHotkey
        );

        void GetShowCmd(
            out uint piShowCmd
        );

        void SetShowCmd(
            uint piShowCmd
        );

        void GetIconLocation(
            [Out(), MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath,
            int cchIconPath,
            out int piIcon
        );

        void SetIconLocation(
            [MarshalAs(UnmanagedType.LPWStr)] string pszIconPath,
            int iIcon
        );

        void SetRelativePath(
            [MarshalAs(UnmanagedType.LPWStr)] string pszPathRel,
            uint dwReserved
        );

        void Resolve(
            IntPtr hWnd,
            uint fFlags
        );

        void SetPath(
            [MarshalAs(UnmanagedType.LPWStr)] string pszFile
        );
    }
}
