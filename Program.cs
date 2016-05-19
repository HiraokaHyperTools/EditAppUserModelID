using EditAppUserModelID.Properties;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace EditAppUserModelID {
    class Program {
        static PropertyKey AppUserModelIDKey = new PropertyKey {
            fmtid = new Guid("{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}"),
            pid = 5,
        };

        /// <summary>
        /// https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa380072(v=vs.85).aspx
        /// </summary>
        static readonly ushort VT_BSTR = 8;
        static readonly ushort VT_LPWSTR = 31;

        /// <summary>
        /// https://msdn.microsoft.com/ja-jp/library/windows/desktop/aa380337(v=vs.85).aspx
        /// </summary>
        static uint STGM_READ = 0;
        static uint STGM_READWRITE = 2;

        static void Main(string[] args) {
            if (args.Length == 2 && args[0] == "/get") {
                var fp = args[1];
                IShellLinkW plnk = (IShellLinkW)new CShellLink();
                IPersistFile ppf = (IPersistFile)plnk;
                ppf.Load(fp, STGM_READ);
                // http://8thway.blogspot.jp/2012/11/csharp-appusermodelid.html
                IPropertyStore pps = (IPropertyStore)plnk;
                PropVariant v = new PropVariant();
                pps.GetValue(ref AppUserModelIDKey, ref v);

                String s = "";
                if (v.vt == VT_BSTR) s = Marshal.PtrToStringUni(v.pVal);
                if (v.vt == VT_LPWSTR) s = Marshal.PtrToStringUni(v.pVal);

                Console.WriteLine(s);
            }
            else if (args.Length == 2 && args[0] == "/list") {
                var fp = args[1];
                IShellLinkW plnk = (IShellLinkW)new CShellLink();
                IPersistFile ppf = (IPersistFile)plnk;
                ppf.Load(fp, STGM_READ);
                IPropertyStore pps = (IPropertyStore)plnk;
                uint cx;
                pps.GetCount(out cx);
                for (uint x = 0; x < cx; x++) {
                    PropertyKey k = new PropertyKey();
                    pps.GetAt(x, ref k);
                    PropVariant v = new PropVariant();
                    pps.GetValue(ref k, ref v);
                    String s = "";
                    if (v.vt == VT_BSTR) s = Marshal.PtrToStringUni(v.pVal);
                    if (v.vt == VT_LPWSTR) s = Marshal.PtrToStringUni(v.pVal);
                    Console.WriteLine(k.fmtid.ToString("B") + " " + k.pid + " " + v.vt + " " + s);
                }
            }
            else if (args.Length == 3 && args[0] == "/set") {
                var fp = args[1];
                var s = args[2];
                IShellLinkW plnk = (IShellLinkW)new CShellLink();
                IPersistFile ppf = (IPersistFile)plnk;
                ppf.Load(fp, STGM_READWRITE);
                // http://8thway.blogspot.jp/2012/11/csharp-appusermodelid.html
                IPropertyStore pps = (IPropertyStore)plnk;
                PropVariant pv = new PropVariant {
                    vt = VT_LPWSTR,
                    pVal = Marshal.StringToBSTR(s)
                };
                pps.SetValue(ref AppUserModelIDKey, ref pv);
                pps.Commit();
                ppf.Save(fp, false);
            }
            else {
                helpYa();
            }
        }

        static void helpYa() {
            Console.Error.WriteLine(Resources.Usage);
            Environment.ExitCode = 1;
        }
    }
}
