using System;
using System.Runtime.InteropServices;
using System.Text;

namespace MiniEXEL.Services
{
    // Resolves the target path of a .lnk shortcut using the native IShellLink COM
    // interface (built into Windows — no extra dependency needed).
    internal static class ShortcutResolver
    {
        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink { }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLinkW
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, IntPtr pfd, int fFlags);
            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
            void SetDescription(string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
            void SetWorkingDirectory(string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
            void SetArguments(string pszArgs);
            void GetHotkey(out short pwHotkey);
            void SetHotkey(short wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
            void SetIconLocation(string pszIconPath, int iIcon);
            void SetRelativePath(string pszPathRel, int dwReserved);
            void Resolve(IntPtr hwnd, int fFlags);
            void SetPath(string pszFile);
        }

        [ComImport]
        [Guid("0000010b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistFile
        {
            void GetClassID(out Guid pClassID);
            void IsDirty();
            void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, int dwMode);
            void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);
            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
            void GetCurFile(out string ppszFileName);
        }

        public static string? ResolveTarget(string lnkPath)
        {
            try
            {
                var link = (IShellLinkW)new ShellLink();
                var file = (IPersistFile)link;
                file.Load(lnkPath, 0); // 0 = read-only

                var sb = new StringBuilder(260);
                link.GetPath(sb, sb.Capacity, IntPtr.Zero, 0);
                var target = sb.ToString();

                Marshal.ReleaseComObject(link);
                return string.IsNullOrWhiteSpace(target) ? null : target;
            }
            catch
            {
                return null;
            }
        }
    }
}
