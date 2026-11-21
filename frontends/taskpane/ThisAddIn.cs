// DataHelper 1.0 VSTO Task Pane Host. Core business rules reside in audited VBA modules.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using Microsoft.Win32;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace DataHelperTaskPanePoc
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObject(0)]
    public sealed class ThisAddIn : AddInBase
    {
        readonly Microsoft.Office.Tools.Excel.ApplicationFactory factory;
        Excel.Application app;
        internal CustomTaskPaneCollection CustomTaskPanes;
        readonly List<PaneRecord> records = new List<PaneRecord>();
        PocAutomation automation;
        string root;
        bool enabled,productMode,autoOpened;
        System.Windows.Forms.Control dispatcher;
        readonly object logGate = new object();
        int uiThread;
        public ThisAddIn(Microsoft.Office.Tools.Excel.ApplicationFactory f, IServiceProvider s)
            : base(f, s, "AddIn", "ThisAddIn") { factory = f; }
        protected override void Initialize()
        {
            base.Initialize();
            app = GetHostItem<Excel.Application>(typeof(Excel.Application), "Application");
            CustomTaskPanes = factory.CreateCustomTaskPaneCollection(null, null, "CustomTaskPanes", "CustomTaskPanes", this);
            automation = new PocAutomation(this);
        }
        protected override void OnStartup()
        {
            base.OnStartup();
            root = Environment.GetEnvironmentVariable("DH_TASKPANE_POC_SESSION");
            enabled = !String.IsNullOrEmpty(root) && File.Exists(Path.Combine(root,"owned-session.txt"));
            if(!enabled) enabled=InitializeInstalledProduct();
            if (!enabled) return;
            uiThread=System.Threading.Thread.CurrentThread.ManagedThreadId;
            dispatcher=new System.Windows.Forms.Control();
            var dispatcherHandle=dispatcher.Handle;
            app.SheetActivate+=OnSheet;
            app.SheetDeactivate+=OnSheet;
            app.WorkbookActivate+=OnWorkbook;
            app.WorkbookDeactivate+=OnWorkbook;
            app.SheetSelectionChange+=OnSelection;
            app.SheetChange+=OnData;
            app.SheetCalculate+=OnSheet;
            Log("vsto-startup", new { pid=Process.GetCurrentProcess().Id, domain=AppDomain.CurrentDomain.FriendlyName,
                addin=GetType().BaseType.FullName, panes=CustomTaskPanes.GetType().FullName, excel=app.Version });
            QueueInitialPane();
        }
        bool InitializeInstalledProduct()
        {
            string frontend=Environment.GetEnvironmentVariable("DH1_FRONTEND");
            if(String.IsNullOrEmpty(frontend))using(var settings=Registry.CurrentUser.OpenSubKey(@"Software\DataHelper\1.0"))
                frontend=settings==null?null:settings.GetValue("Frontend") as string;
            if(frontend!="taskpane")return false;
            // VSTO shadow-copies managed assemblies. Assembly.Location is not
            // a reliable installation directory (the isolated PoC did not expose this).
            string installation;
            using(var settings=Registry.CurrentUser.OpenSubKey(@"Software\DataHelper\1.0"))
                installation=settings==null?null:settings.GetValue("InstallRoot") as string;
            if(String.IsNullOrEmpty(installation))throw new InvalidOperationException("电子羊表格安装目录缺失，请重新运行安装器");
            installation=Path.GetFullPath(installation);
            if(!File.Exists(Path.Combine(installation,"product.json")))throw new InvalidOperationException("电子羊表格产品标记缺失");
            var folder=Path.Combine(installation,"TaskPane");
            // Do not combine a stale cached host with a different on-disk host.
            using(var sha=System.Security.Cryptography.SHA256.Create()){
                var loadedHash=sha.ComputeHash(File.ReadAllBytes(System.Reflection.Assembly.GetExecutingAssembly().Location));
                var installed=sha.ComputeHash(File.ReadAllBytes(Path.Combine(folder,"DataHelperTaskPanePoc.dll")));
                if(Convert.ToBase64String(loadedHash)!=Convert.ToBase64String(installed))throw new InvalidOperationException("电子羊表格宿主缓存不一致，请关闭Excel后修复安装");
            }
            var engine=Path.Combine(installation,"Core","DataHelper_PocEngine.xlam");
            if(!File.Exists(engine))throw new FileNotFoundException("电子羊表格共享引擎缺失",engine);
            bool loaded=false;
            foreach(Excel.Workbook book in app.Workbooks){
                if(book.Name!="DataHelper_PocEngine.xlam")continue;
                if(!String.Equals(book.FullName,engine,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("已加载另一份电子羊表格核心，请关闭Excel后再切换版本");
                loaded=true;
            }
            if(!loaded)app.Workbooks.Open(engine,0,false);
            root=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"DataHelper","Runtime",Process.GetCurrentProcess().Id+"_"+Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(root,"web"));
            foreach(var name in new[]{"index.html","bridge-client.js"})File.Copy(Path.Combine(folder,name),Path.Combine(root,"web",name));
            productMode=true;
            var current=root;
            System.Threading.Tasks.Task.Run(()=>CleanStaleRuntime(Path.GetDirectoryName(current),current));
            return true;
        }
        // N-1: every Excel start creates Runtime\<pid>_<guid>; nothing removed them.
        // Delete only folders of our own pattern whose Excel is gone. When unsure
        // (no access to the process), keep the folder.
        void CleanStaleRuntime(string runtimeRoot,string current)
        {
            int removed=0,kept=0,failed=0;
            try{
                foreach(var dir in Directory.GetDirectories(runtimeRoot)){
                    if(String.Equals(dir,current,StringComparison.OrdinalIgnoreCase))continue;
                    var name=Path.GetFileName(dir);int cut=name.IndexOf('_'),pid;
                    if(cut<=0||name.Length-cut-1!=32||!Int32.TryParse(name.Substring(0,cut),out pid))continue;
                    if(RuntimeOwnerAlive(pid,Directory.GetCreationTime(dir))){kept++;continue;}
                    try{Directory.Delete(dir,true);removed++;}catch{failed++;}
                }
            }catch(Exception error){failed++;try{Log("runtime-cleanup-error",error.Message);}catch{}}
            try{Log("runtime-cleanup",new{removed,kept,failed});}catch{}
        }
        static bool RuntimeOwnerAlive(int pid,DateTime created)
        {
            try{
                using(var process=Process.GetProcessById(pid)){
                    if(!String.Equals(process.ProcessName,"EXCEL",StringComparison.OrdinalIgnoreCase))return false;
                    return process.StartTime<=created.AddSeconds(5);
                }
            }
            catch(ArgumentException){return false;}
            catch{return true;}
        }
        void QueueInitialPane()
        {
            if(!productMode||autoOpened||dispatcher==null)return;
            dispatcher.BeginInvoke(new Action(()=>{
                if(autoOpened||!enabled||app.ActiveWindow==null)return;
                try{Current(true);autoOpened=true;}catch(Exception error){Log("pane-auto-open-error",error.ToString());}
            }));
        }
        protected override void FinishInitialization()
        {
            // The standard VSTO generated host completes startup at this seam.
            OnStartup();
        }
        protected override void InitializeDataBindings()
        {
            BeginInit();
            CustomTaskPanes.BeginInit();
            CustomTaskPanes.EndInit();
            EndInit();
        }
        protected override object RequestService(Guid serviceGuid)
        {
            if (serviceGuid == typeof(Office.IRibbonExtensibility).GUID) return new PocRibbon(this);
            return base.RequestService(serviceGuid);
        }
        protected override object RequestComAddInAutomationService() { return automation; }
        protected override void OnShutdown()
        {
            if (enabled)
            {
                Invalidate("VSTO shutdown");
                app.SheetActivate-=OnSheet;app.SheetDeactivate-=OnSheet;
                app.WorkbookActivate-=OnWorkbook;app.WorkbookDeactivate-=OnWorkbook;
                app.SheetSelectionChange-=OnSelection;app.SheetChange-=OnData;app.SheetCalculate-=OnSheet;
                foreach(var record in records) record.Control.Dispose();
                CustomTaskPanes.Dispose();
                Log("vsto-shutdown", new { pid=Process.GetCurrentProcess().Id, controlsDisposed=records.Count });
                records.Clear();
                dispatcher.Dispose();
                enabled=false;dispatcher=null;
            }
            base.OnShutdown();
        }
        internal void Log(string action, object value)
        {
            if (String.IsNullOrEmpty(root)) return;
            // Runtime cleanup logs from a worker thread; serialise appends.
            lock (logGate) File.AppendAllText(Path.Combine(root,"host-events.jsonl"),new JavaScriptSerializer().Serialize(new {
                time=DateTime.Now.ToString("O"), pid=Process.GetCurrentProcess().Id, action, value })+Environment.NewLine);
        }
        void Invalidate(string reason)
        {
            if (!enabled) return;
            try { app.Run("'DataHelper_PocEngine.xlam'!DHPoc2_Invalidate", reason); } catch { }
            foreach(var record in records) record.Control.NotifyStale(reason);
            Log("state-invalidated",reason);
        }
        void OnSheet(object sheet){Invalidate("worksheet context/data changed");}
        void OnWorkbook(Excel.Workbook book){Invalidate("workbook changed");QueueInitialPane();}
        void OnData(object sheet,Excel.Range range){Invalidate("source data changed");}
        void OnSelection(object sheet,Excel.Range range)
        {
            try
            {
                bool changed=(bool)app.Run("'DataHelper_PocEngine.xlam'!DHPoc2_SelectionEvent",sheet,range);
                Log("selection-event",new {address=range.Address,invalidated=changed});
                if(changed)foreach(var record in records)record.Control.NotifyStale("selection changed");
            }
            catch { }
        }
        PaneRecord Current(bool create)
        {
            PruneClosedWindows();
            Log("pane-thread",new {thread=System.Threading.Thread.CurrentThread.ManagedThreadId,uiThread,apartment=System.Threading.Thread.CurrentThread.GetApartmentState().ToString()});
            if (!enabled || app.ActiveWindow == null) throw new InvalidOperationException("No owned PoC Excel window");
            int hwnd = app.ActiveWindow.Hwnd;
            foreach(var record in records) if(record.Hwnd==hwnd) return record;
            if(!create) return null;
            var control = new PocControl((object)app, root, Log);
            var pane = CustomTaskPanes.Add(control,"电子羊表格",app.ActiveWindow);
            pane.DockPosition=Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.Width=PaneWidth(hwnd);
            var entry = new PaneRecord { Hwnd=hwnd, Pane=pane, Control=control };
            pane.VisibleChanged += (s,e) => { Log("pane-visible",new {hwnd,visible=pane.Visible}); if(!pane.Visible) Invalidate("pane closed"); };
            records.Add(entry);
            pane.Visible=true;
            control.BeginInvoke(new Action(async () => { try { await control.Start(); } catch(Exception error) { Log("webview-start-error",error.ToString()); } }));
            Log("pane-created", new {hwnd,dock=pane.DockPosition.ToString(),width=pane.Width});
            return entry;
        }
        // The layout is designed for 420 logical px. CustomTaskPane.Width is in device
        // pixels (150% DPI: 420 -> 270 CSS px), so scale it, but keep it within 40%
        // of the screen so small or high-DPI laptops still see the sheet.
        static int PaneWidth(int hwnd)
        {
            try{
                float dpi;
                using(var g=System.Drawing.Graphics.FromHwnd(new IntPtr(hwnd)))dpi=g.DpiX;
                int wanted=(int)Math.Round(420*dpi/96f);
                int cap=(int)(System.Windows.Forms.Screen.FromHandle(new IntPtr(hwnd)).WorkingArea.Width*0.4);
                return FitPaneWidth(wanted,cap);
            }catch{return 420;}
        }
        // Cap wins on narrow screens. The HTML remains scrollable below its ideal width.
        internal static int FitPaneWidth(int wanted,int cap)
        {
            return Math.Max(1,Math.Min(Math.Max(420,wanted),cap));
        }
        public void Toggle()
        {
            var record=Current(false);
            if(record==null) Current(true); else record.Pane.Visible=!record.Pane.Visible;
            Log("ribbon-toggle",true);
        }
        public void SetVisible(bool visible)
        {
            var record=Current(visible);
            if(record!=null) record.Pane.Visible=visible;
        }
        public string State()
        {
            try{PruneClosedWindows();}catch(Exception error){Log("state-prune-error",error.ToString());return new JavaScriptSerializer().Serialize(new{enabled,pid=Process.GetCurrentProcess().Id,stateError=error.Message,panes=new object[0]});}
            var recordsDto = new List<object>();
            foreach(var record in records){
                try{
                    if(record.Pane==null||record.Control==null)throw new InvalidOperationException("Pane registration is incomplete");
                    recordsDto.Add(new {hwnd=record.Hwnd, visible=record.Pane.Visible,
                        dock=record.Pane.DockPosition.ToString(),width=record.Pane.Width,height=record.Pane.Height,
                        ready=record.Control.IsUsable,browserPid=record.Control.BrowserPid,reply=record.Control.LastReply});
                }catch(Exception error){
                    Log("state-pane-error",new{hwnd=record.Hwnd,detail=error.ToString()});
                    recordsDto.Add(new{hwnd=record.Hwnd,ready=false,stateError=error.Message});
                }
            }
            return new JavaScriptSerializer().Serialize(new {enabled,pid=Process.GetCurrentProcess().Id,panes=recordsDto});
        }
        void PruneClosedWindows()
        {
            // Excel destroys a window's native task pane when that window closes.
            // Keep no stale COM pane/control in this disposable host's registry.
            var live=new HashSet<int>();
            foreach(Excel.Window window in app.Windows) if(window!=null)live.Add(window.Hwnd);
            // Excel can temporarily expose an empty Windows collection while a
            // valid ActiveWindow is still present (observed in a real SDI host).
            // An incomplete COM inventory is not proof that an owned window closed.
            var active=app.ActiveWindow;
            if(active!=null)live.Add(active.Hwnd);
            bool incomplete=app.Windows.Count==0;
            for(int i=records.Count-1;i>=0;i--)
            {
                var record=records[i];
                if(live.Contains(record.Hwnd))continue;
                if(incomplete&&IsWindow(new IntPtr(record.Hwnd)))continue;
                records.RemoveAt(i);
                record.Control.Dispose();
                try { CustomTaskPanes.Remove(record.Pane); }
                catch(ArgumentException) { }
                catch(InvalidOperationException) { }
                catch(COMException) { }
                Log("closed-window-pruned",new {hwnd=record.Hwnd});
            }
        }
        [DllImport("user32.dll")]
        static extern bool IsWindow(IntPtr hwnd);
        public void Trigger(string command) { Current(true).Control.Trigger(command); }
        public void Request(string wire) { Current(true).Control.Submit(wire); }
        public void Dispatch(Action action)
        {
            if(dispatcher==null)throw new InvalidOperationException("Owned VSTO dispatcher not initialized");
            if(dispatcher.InvokeRequired)dispatcher.Invoke(action);else action();
        }
        public string DispatchState()
        {
            if(dispatcher==null||dispatcher.IsDisposed)return "{\"enabled\":false,\"stateError\":\"HOST_STOPPED\"}";
            try{return dispatcher.InvokeRequired?(string)dispatcher.Invoke(new Func<string>(State)):State();}
            catch(Exception error){Log("state-dispatch-error",error.ToString());return new JavaScriptSerializer().Serialize(new{enabled,pid=Process.GetCurrentProcess().Id,stateError=error.Message,panes=new object[0]});}
        }
        sealed class PaneRecord { public int Hwnd; public CustomTaskPane Pane; public PocControl Control; }
    }
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class PocAutomation
    {
        readonly ThisAddIn host;
        public PocAutomation(ThisAddIn host) { this.host=host; }
        public string StateJson() { return host.DispatchState(); }
        public void ShowPane() { host.Dispatch(()=>host.SetVisible(true)); }
        public void HidePane() { host.Dispatch(()=>host.SetVisible(false)); }
        public void Trigger(string command) { host.Dispatch(()=>host.Trigger(command)); }
        // Isolated test host only. Never a production JS bridge or business API.
        public void Request(string wire) { host.Dispatch(()=>host.Request(wire)); }
    }
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PocRibbon : Office.IRibbonExtensibility
    {
        readonly ThisAddIn host;
        public PocRibbon(ThisAddIn host) { this.host=host; }
        public string GetCustomUI(string id) { return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs><tab id=\"dhPocTab\" label=\"电子羊表格\"><group id=\"dhPocGroup\" label=\"工作台\"><button id=\"dhPocToggle\" label=\"打开/关闭电子羊表格\" size=\"large\" imageMso=\"TaskPaneInsert\" onAction=\"Toggle\"/></group></tab></tabs></ribbon></customUI>"; }
        public void Toggle(Office.IRibbonControl control) { host.Toggle(); }
    }
}
