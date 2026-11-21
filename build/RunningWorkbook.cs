using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class RunningWriteWorkbook
{
    [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved,out IRunningObjectTable table);
    [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved,out IBindCtx context);
    public static object Find(string fullPath)
    {
        IRunningObjectTable table=null;IBindCtx context=null;IEnumMoniker enumerator=null;
        try{
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0,out table));Marshal.ThrowExceptionForHR(CreateBindCtx(0,out context));
            table.EnumRunning(out enumerator);var current=new IMoniker[1];
            while(enumerator.Next(1,current,IntPtr.Zero)==0){
                try{string name;current[0].GetDisplayName(context,null,out name);
                    if(String.Equals(name,fullPath,StringComparison.OrdinalIgnoreCase)){object workbook;table.GetObject(current[0],out workbook);return workbook;}
                }catch(COMException){}finally{Marshal.ReleaseComObject(current[0]);}
            }
            return null;
        }finally{
            if(enumerator!=null)Marshal.ReleaseComObject(enumerator);if(context!=null)Marshal.ReleaseComObject(context);if(table!=null)Marshal.ReleaseComObject(table);
        }
    }
}

// Unicode shell-link writer. WScript shortcut Save corrupts CJK filenames on some ACP settings.
public static class DHDesktopShortcut
{
    [ComImport, Guid("00021401-0000-0000-C000-000000000046")] private class ShellLinkObject { }
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
    private interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder path, int capacity, IntPtr data, uint flags);
        void GetIDList(out IntPtr list); void SetIDList(IntPtr list);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder text, int capacity);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string text);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder text, int capacity);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string text);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder text, int capacity);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string text);
        void GetHotkey(out short hotkey); void SetHotkey(short hotkey);
        void GetShowCmd(out int command); void SetShowCmd(int command);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder text, int capacity, out int index);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string path, int index);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string path, uint reserved);
        void Resolve(IntPtr window, uint flags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string path);
    }
    public static void Write(string file, string target, string arguments, string description)
    { Write(file, target, arguments, description, null); }
    // icon: optional .ico shown instead of the PowerShell target's icon.
    public static void Write(string file, string target, string arguments, string description, string icon)
    {
        object instance = new ShellLinkObject();
        try {
            var link = (IShellLinkW)instance;
            link.SetPath(target); link.SetArguments(arguments); link.SetDescription(description); link.SetShowCmd(7);
            if (!String.IsNullOrEmpty(icon) && System.IO.File.Exists(icon)) link.SetIconLocation(icon, 0);
            ((IPersistFile)instance).Save(file, true);
        } finally { Marshal.FinalReleaseComObject(instance); }
    }
    public static string[] Read(string file)
    {
        object instance = new ShellLinkObject();
        try {
            ((IPersistFile)instance).Load(file, 0); var link=(IShellLinkW)instance;
            var target=new System.Text.StringBuilder(32768); var args=new System.Text.StringBuilder(32768);
            link.GetPath(target,target.Capacity,IntPtr.Zero,4); link.GetArguments(args,args.Capacity);
            return new string[]{target.ToString(),args.ToString()};
        } finally { Marshal.FinalReleaseComObject(instance); }
    }
}