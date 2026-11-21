using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
namespace DataHelperTaskPanePoc
{
    public sealed class PocControl : UserControl
    {
        readonly WebView2 view=new WebView2();readonly dynamic excel;readonly string root;
        readonly Action<string,object> log;readonly JavaScriptSerializer json=new JavaScriptSerializer();
        readonly HashSet<string> requests=new HashSet<string>();
        // The page asks for context while it loads, before Ready (R15-RT-05).
        // Replies produced then are held and sent once the view is usable.
        readonly List<string> heldReplies=new List<string>();
        const string Page="https://datahelper-taskpane-poc.local/index.html";
        bool busy,released;uint browserPid;public bool Ready{get;private set;}public string LastReply{get;private set;}
        // A disposed/failed WebView can have no CoreWebView2 even after navigation
        // once completed. Observation must not dereference that lifecycle gap.
        public uint BrowserPid{get{return !released&&Ready?browserPid:0;}}
        public bool IsUsable{get{return !released&&!IsDisposed&&Ready&&view.CoreWebView2!=null;}}
        public PocControl(object application,string folder,Action<string,object> logger)
        {excel=application;root=folder;log=logger;Dock=DockStyle.Fill;view.Dock=DockStyle.Fill;Controls.Add(view);}
        public async Task Start()
        {
            // One stable profile for every Excel process (R15-RT-03): the per-process
            // runtime folder made the receipt's localStorage vanish on restart.
            var env=await CoreWebView2Environment.CreateAsync(null,Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"DataHelper","WebProfile"));
            await view.EnsureCoreWebView2Async(env);
            view.CoreWebView2.Settings.AreHostObjectsAllowed=false;view.CoreWebView2.Settings.AreDefaultContextMenusEnabled=false;
            view.CoreWebView2.NavigationStarting+=(s,e)=>{if(e.Uri!=Page)e.Cancel=true;};
            view.CoreWebView2.NewWindowRequested+=(s,e)=>e.Handled=true;
            view.CoreWebView2.SetVirtualHostNameToFolderMapping("datahelper-taskpane-poc.local",Path.Combine(root,"web"),CoreWebView2HostResourceAccessKind.DenyCors);
            view.CoreWebView2.WebMessageReceived+=Receive;
            var loaded=new TaskCompletionSource<bool>();view.CoreWebView2.NavigationCompleted+=(s,e)=>loaded.TrySetResult(e.IsSuccess);
            view.CoreWebView2.Navigate(Page);
            if(await Task.WhenAny(loaded.Task,Task.Delay(15000))!=loaded.Task||!loaded.Task.Result)throw new Exception("Local page failed to load");
            if(released||IsDisposed)throw new ObjectDisposedException("电子羊表格窗格在启动期间已关闭");
            browserPid=view.CoreWebView2.BrowserProcessId;
            Ready=true;log("webview-ready",new{runtime=env.BrowserVersionString,browserPid=BrowserPid,hostPid=Process.GetCurrentProcess().Id});
            if(heldReplies.Count>0){log("held-replies-flushed",new{count=heldReplies.Count});foreach(var held in heldReplies)view.CoreWebView2.PostWebMessageAsJson(held);heldReplies.Clear();}
        }
        object Normalize(object value)
        {
            if(value==null||value is DBNull)return null;var array=value as Array;
            if(array!=null){var list=new List<object>();foreach(object item in array)list.Add(Normalize(item));return list;}return value;
        }
        object Run(string operation,object argument=null)
        {
            log("vba-enter",new{operation,thread=System.Threading.Thread.CurrentThread.ManagedThreadId});
            var envelope=(IList)Normalize(argument==null?excel.Run("'DataHelper_PocEngine.xlam'!DHPoc2_Invoke",operation):excel.Run("'DataHelper_PocEngine.xlam'!DHPoc2_Invoke",operation,argument));
            log("vba-return",new{operation});
            if((string)envelope[0]!="ok")throw new InvalidOperationException("VBA ["+envelope[1]+"] "+envelope[2]);return envelope[1];
        }
        void Receive(object sender,CoreWebView2WebMessageReceivedEventArgs e)
        {
            if(e.Source!=Page||released)return;BridgeRequest r=null;
            try{
                r=BridgeRequest.Parse(e.WebMessageAsJson);
                if(requests.Count>=10000)throw new InvalidOperationException("Request budget exceeded; reopen pane");
                if(!requests.Add(r.Id))throw new InvalidOperationException("Request replay rejected");
                if(busy)throw new InvalidOperationException("BUSY: native confirmation or execution pending");
                busy=true;log("request",new{r.Id,r.Command,r.Actions});
                if(r.Command=="confirm")BeginInvoke(new Action(()=>ConfirmNextTurn(r)));
                else BeginInvoke(new Action(()=>ProcessRequest(r)));
            }catch(Exception error){Reply(new{v=1,id=r==null?null:r.Id,command=r==null?"invalid-message":r.Command,ok=false,error=error.Message});}
        }
        void ProcessRequest(BridgeRequest r)
        {
            try{
                object result;
                switch(r.Command){
                    case "context":result=Run("DHPoc2_Context");break;
                    case "candidate":result=Run("DHPoc2_Candidate");break;
                    case "scan":result=Run("DHPoc2_Scan");break;
                    case "workbook":result=Run("DHPoc2_Workbook");break;
                    case "license-status":result=Run("DHPoc2_LicenseStatus");break;
                    case "license-activate":result=Run("DHPoc2_LicenseActivate",r.License);break;
                    case "lookup_meta":result=Run("DHPoc2_LookupMeta");break;
                    case "lookup_preview":result=Run("DHPoc2_LookupPreview",r.Actions);break;
                    case "append_preview":result=Run("DHPoc2_AppendPreview",r.Actions);break;
                    case "unpivot_preview":result=Run("DHPoc2_UnpivotPreview",r.Actions);break;
                    case "split_sheet_preview":result=Run("DHPoc2_SplitSheetPreview",r.Actions);break;
                    case "prepare":result=Run("DHPoc2_Prepare",r.Actions);break;
                    case "invalidated":Run("DHPoc2_Invalidate","client explicitly discarded result");result="INVALIDATED";break;
                    case "execute":log("execution-request",new{r.Id});result=Run("DHPoc2_Execute");break;
                    case "telemetry":result=new{width=Width,height=Height};break;
                    case "select-range":{
                        var addr=Convert.ToString(r.Path);
                        if(!Regex.IsMatch(addr,@"\A\$?[A-Z]{1,3}\$?[0-9]{1,7}(:\$?[A-Z]{1,3}\$?[0-9]{1,7})?\z"))throw new InvalidOperationException("地址不合法");
                        Run("DHPoc2_SelectRange",addr);
                        result="SELECTED";break;}
                    case "open-path":{
                        var target=Convert.ToString(r.Path);
                        var root=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ExcelDataHelper");
                        var full=System.IO.Path.GetFullPath(target);
                        if(!full.StartsWith(root,StringComparison.OrdinalIgnoreCase)||!File.Exists(full))throw new InvalidOperationException("只允许打开本工具备份目录里的文件");
                        Process.Start("explorer.exe","/select,\""+full+"\"");
                        result="OPENED";break;}
                    default:throw new InvalidOperationException("Unmapped command");
                }
                bool success=true;if(r.Command=="execute"){var row=(IList)result;success=(string)row[0]=="COMPLETED";}
                Reply(new{v=1,id=r.Id,command=r.Command,ok=success,result,error=success?null:"Execution did not complete; inspect phase, audit and backup in result"});
            }catch(Exception error){Reply(new{v=1,id=r==null?null:r.Id,command=r==null?"invalid-message":r.Command,ok=false,error=error.Message});}
            finally{busy=false;}
        }
        void ConfirmNextTurn(BridgeRequest r)
        {
            try{
                var target=(IList)Run("DHPoc2_ConfirmationTarget");
                // R15: readable plan (target[5]) from the engine; older engines fall back to raw action ids.
                var plan=target.Count>5?Convert.ToString(target[5]):("操作："+target[4]+"\n");
                // 弱保软：主按钮只问一句「要不要整理」，步骤名压成一行；细节留在面板上，不再糊一屏
                var stepNames=new List<string>();
                foreach(var raw in (plan??"").Split(new[]{'\n'},StringSplitOptions.RemoveEmptyEntries)){
                    var t=raw.Trim();
                    var dot=t.IndexOf('.');
                    if(t.Length>2&&char.IsDigit(t[0])&&dot>0)stepNames.Add(t.Substring(dot+1).Trim());
                }
                var brief=stepNames.Count>0?string.Join("、",stepNames):"把这张表整理一下";
                var answer=MessageBox.Show(this,
                    "要帮你整理「"+target[2]+" "+target[3]+"」吗？\n\n会做："+brief+"\n\n会先备份，文件不会自动保存。",
                    "电子羊表格",MessageBoxButtons.YesNo,MessageBoxIcon.Question,MessageBoxDefaultButton.Button1);
                if(answer==DialogResult.Yes)Reply(new{v=1,id=r.Id,command="confirm",ok=true,result=Run("DHPoc2_Confirm",target[0])});
                else{Run("DHPoc2_Invalidate","native confirmation cancelled");Reply(new{v=1,id=r.Id,command="confirm",ok=false,error="CANCELLED: no write authorised"});}
            }catch(Exception error){Reply(new{v=1,id=r.Id,command="confirm",ok=false,error=error.Message});}
            finally{busy=false;}
        }
        public void NotifyStale(string reason)
        {if(IsUsable)view.CoreWebView2.PostWebMessageAsJson(json.Serialize(new{v=1,type="invalidated",command="invalidated",reason}));}
        void Reply(object reply)
        {
            LastReply=json.Serialize(reply);
            try{log("reply",reply);}catch(Exception error){LastReply=json.Serialize(new{v=1,ok=false,error="HOST_LOG_FAILED: result may have been written; inspect VBA log/backup: "+error.Message,result=reply});}
            if(IsUsable)view.CoreWebView2.PostWebMessageAsJson(LastReply);
            else if(!released&&!Ready&&heldReplies.Count<20){heldReplies.Add(LastReply);log("reply-held",new{count=heldReplies.Count});}
        }
        // Typed replacement for EvalForTest; sends data through the real WebView channel.
        public async void Submit(string wire)
        {
            try{
                if(!IsUsable){
                    var parsed=json.DeserializeObject(wire) as Dictionary<string,object>;
                    Reply(new{v=1,id=parsed!=null&&parsed.ContainsKey("id")?parsed["id"]:null,command="not-ready",ok=false,error="NOT_READY: wait for this Excel window pane"});return;
                }
                await view.ExecuteScriptAsync("window.chrome.webview.postMessage(JSON.parse("+json.Serialize(wire)+"))");
            }catch(Exception error){Reply(new{v=1,command="submit-error",ok=false,error.Message});}
        }
        public void Trigger(string command){Submit(json.Serialize(new{v=1,id="host_"+Guid.NewGuid().ToString("N"),command}));}
        protected override void Dispose(bool disposing)
        {if(disposing&&!released){uint browser=browserPid;released=true;Ready=false;view.Dispose();log("webview-disposed",new{browserPid=browser});}base.Dispose(disposing);}
    }
}
