using System;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Reflection;
[assembly: AssemblyTitle("电子羊表格安装程序")]
[assembly: AssemblyProduct("电子羊表格")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
internal static class Setup
{
 internal static string FailureMessage(string output,string error){
  const string prefix="DH_SETUP_EXPECTED_ERROR:";
  using(var lines=new StringReader((error??"")+"\n"+(output??""))){
   string line;while((line=lines.ReadLine())!=null){
    if(line.StartsWith(prefix,StringComparison.Ordinal))return line.Substring(prefix.Length).Trim()+"\n\n安装未完成，现有安装未被本次预检改动。详细信息已写入本机安装日志目录。";
   }
  }
  return (error??"")+"\n安装未完成。详细信息已写入本机安装日志目录。";
 }
 [STAThread] static int Main(string[] args){
  string folder=null;
  try{
   byte[] bytes;using(var resource=Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))using(var memory=new MemoryStream()){resource.CopyTo(memory);bytes=memory.ToArray();}
   string hash;using(var sha=SHA256.Create())hash=BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-","");
   if(hash!=BuildInfo.Hash)throw new InvalidDataException("安装资源校验失败。");
   folder=Path.Combine(Path.GetTempPath(),"DataHelperSetup_"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
   using(var stream=new MemoryStream(bytes))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))foreach(var entry in zip.Entries){
    var destination=Path.GetFullPath(Path.Combine(folder,entry.FullName));if(!destination.StartsWith(folder+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("Unsafe package path");
    Directory.CreateDirectory(Path.GetDirectoryName(destination));entry.ExtractToFile(destination);
   }
   bool noLaunch=Array.IndexOf(args,"--no-launch")>=0;
   string command="-NoProfile -ExecutionPolicy Bypass -File \""+Path.Combine(folder,"install.ps1")+"\" -Edition "+BuildInfo.Edition+(noLaunch?" -NoLaunch":"");
   var start=new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),@"WindowsPowerShell\v1.0\powershell.exe"),command){UseShellExecute=false,CreateNoWindow=true,RedirectStandardError=true,RedirectStandardOutput=true,StandardOutputEncoding=new System.Text.UTF8Encoding(false),StandardErrorEncoding=new System.Text.UTF8Encoding(false)};
   using(var process=Process.Start(start)){
    var outputTask=process.StandardOutput.ReadToEndAsync();var errorTask=process.StandardError.ReadToEndAsync();process.WaitForExit();
    string output=outputTask.Result,error=errorTask.Result;
    var logs=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"DataHelper","SetupLogs");Directory.CreateDirectory(logs);
    File.WriteAllText(Path.Combine(logs,DateTime.Now.ToString("yyyyMMdd_HHmmssfff")+"_"+BuildInfo.Edition+".txt"),output+Environment.NewLine+error);
    if(process.ExitCode!=0)throw new Exception(FailureMessage(output,error));
   }
   return 0;
  }catch(Exception error){if(Array.IndexOf(args,"--no-launch")<0&&Array.IndexOf(args,"--quiet-errors")<0)MessageBox.Show(error.Message,"电子羊表格 1.0.0 · 安装",MessageBoxButtons.OK,MessageBoxIcon.Warning);return 1;}
  finally{if(folder!=null&&Directory.Exists(folder))try{Directory.Delete(folder,true);}catch{}}
 }
}
