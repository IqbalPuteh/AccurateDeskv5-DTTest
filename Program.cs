using System;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using System.Runtime.InteropServices;
using FlaUI.Core;
using FlaUI.UIA3;
using FlaUI.Core.Input;
using FlaUI.Core.Conditions;
using FlaUI.Core.Definitions;
using FlaUI.Core.AutomationElements;
using Serilog;
using System.Threading;
using Windows.Media.Capture;
using iPos4DS_DTTest;
using System.IO.Compression;

namespace AccurateDeskv5_DTTest 
{
    internal class Program
    {
        static Application appx;
        static UIA3Automation automationUIA3;
        static ConditionFactory cf = new ConditionFactory(new UIA3PropertyLibrary());
        static AutomationElement window;
        static string dtID = ConfigurationManager.AppSettings["dtID"];
        static string dtName = ConfigurationManager.AppSettings["dtName"];
        static string loginId = ConfigurationManager.AppSettings["loginId"];
        static string password = ConfigurationManager.AppSettings["password"];
        static string erpappnamepath = ConfigurationManager.AppSettings["erpappnamepath"];
        static string dbaliasflag = ConfigurationManager.AppSettings["dbaliasflag"].ToUpper();
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        static string xPosWelcome = ConfigurationManager.AppSettings["xPosWelcome"];
        static string dbaliasname = ConfigurationManager.AppSettings["dbaliasname"];
        static string dbpath = ConfigurationManager.AppSettings["dbpath"];
        static string yPosWelcome = ConfigurationManager.AppSettings["yPosWelcome"];
        static string logfilename = "";
        static int pid = 0;

        [DllImport("user32.dll")]

        public static extern bool BlockInput(bool fBlockIt);

        private static AutomationElement WaitForElement(Func<AutomationElement> findElementFunc)
        {
            AutomationElement element = null;
            for (int i = 0; i < 2000; i++)
            {
                element = findElementFunc();
                if (element is not null)
                {
                    break;
                }

                //Thread.Sleep(1);
            }
            return element;
        }


        static void Main(string[] args)
        {
            try
            {
                //* Call this method to disable keyboard input
                int maxWidth = Console.LargestWindowWidth;
                Console.Title = "Accurate Desktop Version 5 - Automasi - By PT FAIRBANC TECHNOLOGIEST INDONESIA";
                Console.WindowLeft = 0;
                Console.WindowTop = 0;
                Console.SetWindowPosition(0, 0);
                Console.WindowWidth = maxWidth;
                Console.BackgroundColor = ConsoleColor.DarkGray;
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine($"");
                Console.WriteLine($"******************************************************************");
                Console.WriteLine($"                    Automasi akan dimulai !                       ");
                Console.WriteLine($"             Keyboard dan Mouse akan di matikan...                ");
                Console.WriteLine($"     Komputer akan menjalankan oleh applikasi robot automasi...   ");
                Console.WriteLine($" Aktifitas penggunakan komputer akan ter-BLOKIR sekitar 10 menit. ");
                Console.WriteLine($"******************************************************************");

#if DEBUG
                BlockInput(false);
#else
                BlockInput(true);
#endif
                var myFileUtil = new MyDirectoryManipulator();
                if (!Directory.Exists(appfolder))
                {
                    myFileUtil.CreateDirectory(appfolder);
                    myFileUtil.CreateDirectory(uploadfolder);
                    myFileUtil.CreateDirectory(sharingfolder);
                }
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.BackgroundColor = ConsoleColor.Black;
                var temp1 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Excel);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp1}"));
                do
                {
                } while (!Task.CompletedTask.IsCompleted);
                var temp2 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Log);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp2}"));
                var temp3 = myFileUtil.DeleteFiles(appfolder, MyDirectoryManipulator.FileExtension.Zip);
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] {temp3}"));
                var config = new LoggerConfiguration();
                logfilename = "DEBUG-" + dtID + "-" + dtName + ".log";
                config.WriteTo.File(appfolder + Path.DirectorySeparatorChar + logfilename);
                if (enableconsolelog == "Y")
                {
                    config.WriteTo.Console();
                }
                Log.Logger = config.CreateLogger();

                Log.Information("Accurate Desktop ver.5 Automation - *** Started *** ");
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();

                
                if (!OpenAppAndDBConfig())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenAppAndDBConfig) !!!");
                    return;
                }
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();
                if (!LoginApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (LoginApp) !!!");
                    return;
                }
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();
                if (!OpenReport("sales"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenReport -> Sales) !!!");
                    return;
                }
                automationUIA3 = new UIA3Automation();
                window = automationUIA3.GetDesktop();
                if (!ClosingWorkspace())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ClosingWorkspace) !!!");
                    return;
                }
                window = automationUIA3.GetDesktop();
                if (!OpenReport("ar"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenReport -> Sales) !!!");
                    return;
                }
                window = automationUIA3.GetDesktop();
                if (!ClosingWorkspace())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ClosingWorkspace) !!!");
                    return;
                }
                window = automationUIA3.GetDesktop();
                if (!OpenReport("outlet"))
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenReport -> Sales) !!!");
                    return;
                }
                window = automationUIA3.GetDesktop();
                if (!ClosingWorkspace())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (ClosingWorkspace) !!!");
                    return;
                }
                window = automationUIA3.GetDesktop();
                if (!CloseApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (CloseApp) !!!");
                    return;
                }

                ZipandSend();

            }
            catch (Exception ex)
            {
                Log.Information($"iPos v5 automation error => {ex.ToString()}");
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")} INF] iPos automation error => {ex.ToString()}"));
            }
            finally
            {
                Console.Beep();
                Task.Delay(500);
                Console.Beep();
                Task.Delay(500);
                //* Call this method to enable keyboard input
                BlockInput(false);

                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Accurate Desktop ver.5 Automation - ***   END   ***"));
                if (automationUIA3 is not null)
                {
                    automationUIA3.Dispose();
                }
                Log.CloseAndFlush();
            }
        }
        static bool OpenAppAndDBConfig()
        {
            var step = 0;
            try
            {

                //ProcessStartInfo psi = new ProcessStartInfo();
                //psi.FileName = Environment.GetEnvironmentVariable("WINDIR") + "\\explorer.exe";
                //psi.Verb = "runasuser";
                //psi.LoadUserProfile = true;
                //psi.Arguments = @$"{erpappnamepath}";
                //Process p = Process.Start(psi);
                //Log.Information(p.ProcessName);
                //Thread.Sleep(25000);

                /* added on Tue 2 April 2024, by Iqbal, for anticipating standart user windows cannot see process name */
                // Specify the path to your shortcut
                string shortcutPath = $@"{erpappnamepath}";
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = shortcutPath;
                startInfo.UseShellExecute = true;
                startInfo.CreateNoWindow = false;

                Process process = new Process();
                process.StartInfo = startInfo;
                automationUIA3 = new UIA3Automation();

                try
                {
                    //appx = Application.Launch(process.StartInfo);
                    process.Start();
                    pid =  process.Id;
                    Log.Information($"Accurate for automation has a PID# =>> {process.Id}");
                    Thread.Sleep(15000);
                }
                catch (Exception ex)
                {
                    Log.Information("Error =>" + ex.Message);
                    Log.Information($"[OpenAppAndDBConfig] Error ketika menghandle opening ACCURATE (window) process..."); 
                    return false; 
                }

                window = automationUIA3.GetDesktop();
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllDescendants(cf.ByName("ACCURATE 5", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                //if (auEle.Length > 0)
                //{
                //    // Get the process ID of MyUnElevatedProcess.exe
                //    string targetProcessName = "accurate";
                //    Process[] processes = Process.GetProcessesByName(targetProcessName);
                //    if (processes.Length > 0)
                //    {
                //        Process latestProcess = processes.OrderByDescending(p => p.StartTime).First();
                //        pid = latestProcess.Id;
                //    }
                //}
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                step++;
                if (mainElement is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function (no PID found) !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                AutomationElement selamatWindowEle = null;
                var auEle1 = window.FindAllChildren(cf.ByName("Selamat Datang", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle1)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        selamatWindowEle = item;
                    }
                }
                step++;
                if (selamatWindowEle is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function !!");
                    return false;
                }
                var par = selamatWindowEle.BoundingRectangle;
                Int16 cordX = Convert.ToInt16(ConfigurationManager.AppSettings["xPosWelcome"]);
                Int16 cordY = Convert.ToInt16(ConfigurationManager.AppSettings["yPosWelcome"]);
                BlockInput(false);
                Thread.Sleep(1000);
                selamatWindowEle.SetForeground();
                Mouse.MoveTo(cordX, cordY);
                Mouse.Click();
                selamatWindowEle = null;
                Log.Information($"Action -> Closing 'Welcome to Accurate' window at {(cordX)},{(cordY)}.");
                Thread.Sleep(1500);
                BlockInput(true);

                //Logic to handle failed to close 'Selamat data window'
                auEle1 = window.FindAllChildren(cf.ByName("Selamat Datang", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle1)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        selamatWindowEle = item;
                    }
                }
                step++;
                if (selamatWindowEle != null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function. (failed to close 'Selamat Datang' window).");
                    return false;
                }

                AutomationElement ele = null;
                auEle = window.FindAllDescendants(cr => cr.ByClassName("TsuiSkinMenuBar"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele = item;
                        break;
                    }
                }
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function !!");
                    return false;
                }
                Log.Information("Element Interaction on property class named -> " + ele.Properties.ClassName.ToString());
                ele.SetForeground();

                //ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("File")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Berkas")));
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.Click();
                Thread.Sleep(1000);

                // Context
                ele = null;
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                //System.Windows.Forms.SendKeys.SendWait("%F");
                //Log.Information("Sending keys 'ALT+F'...");
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                Thread.Sleep(1000);

                //Using opened Database window
                // ele = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Buka Database")));
                ele = null;
                auEle = window.FindAllDescendants(cr => cr.ByName("Buka Database"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele = item;
                        break;
                    }
                }
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                ele.SetForeground();

                if (dbaliasflag.Trim() == "Y")
                {

                    //* Opening alias DB by clicking 'Alias'' Button
                    ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Alias")));
                    step++;
                    if (ele is null)
                    {
                        Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function -> DBalias flag - true !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.AsButton().Click();
                    Thread.Sleep(1000);

                    //* Findng alias window under desktop
                    var MainEle = WaitForElement(() => window.FindFirstChild(cr => cr.ByName("Alias")));
                    step++;
                    if (ele is null)
                    {
                        Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function -> DBalias flag is true !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + MainEle.Properties.Name.ToString());
                    MainEle.SetForeground();

                    //* clicking on {DBaliasname}
                    ele = WaitForElement(() => MainEle.FindFirstDescendant(cr => cr.ByName(dbaliasname)));
                    step++;
                    if (ele is null)
                    {
                        Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function -> DBalias flag is true !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.Click();

                    //* clicking on 'Buka' button
                    ele = WaitForElement(() => MainEle.FindFirstDescendant(cr => cr.ByName("Buka")));
                    step++;
                    if (ele is null)
                    {
                        Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function -> DBalias flag is true !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                    ele.AsButton().Click();
                    Thread.Sleep(1000);

                }
                else
                {
                    var ele2 = WaitForElement(() => ele.FindFirstChild(cr => cr.ByClassName("TEdit")));
                    step ++;
                    if (ele2 is null)
                    {
                        Log.Information($"[Step #{step}] Quitting, end of OpenAppAndDBConfig function -> DBalias flag is false!!");
                        return false;
                    }
                    Log.Information("Element Interaction on property class named -> " + ele2.Properties.ClassName.ToString());

                    if (ele2.AsTextBox().Text != $@"{dbpath}")
                    {
                        Thread.Sleep(1000);
                        ele2.AsTextBox().Text = $@"{dbpath}";
                    }

                    ele = ele.FindFirstChild(cf => cf.ByName("OK")).AsButton();
                    Log.Information("Clicking 'OK' button...");
                    ele.Click();
                }
                return true;
            }
            catch
            {
                Log.Information("Quitting, end of DB automation function !!");
                return false;
            }
        }
        static bool LoginApp()
        {
            var step = 0;
            try
            {
                Thread.Sleep(5000);
                AutomationElement ele = null;
                AutomationElement[] auEle = window.FindAllDescendants(cr => cr.ByName("Daftar"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele = item;
                        break;
                    }
                }
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of LoginApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                var ele2 = ele.FindFirstDescendant(cf => cf.ByClassName("TEdit")).AsTextBox();
                ele2.Enter(loginId + "\t");
                Log.Information("Sending Login Id...");

                System.Windows.Forms.SendKeys.SendWait(password);
                Log.Information("Sending password...");

                ele.FindFirstDescendant(cf => cf.ByName("OK")).AsButton().Click();
                Log.Information("Clicking 'OK' button...");
                Log.Information("Now wait for 1 minute before clicking report...");
                Thread.Sleep(15000);
                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.ToString());
                return false;
            }
        }

        private static bool OpenReport(string reportType)
        {
            Thread.Sleep(3000);
            var step = 0;
            try
            {
                AutomationElement mainElement = null;
                //AutomationElement[] auEle = window.FindAllChildren(cf.ByName("ACCURATE 5", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                // TfrmMain
                AutomationElement[] auEle = window.FindAllChildren(cf.ByClassName("TfrmMain"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                step += 1;
                if (mainElement is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                step += 1;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                Thread.Sleep(500);

                /// Click on Reports menu ///
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Laporan")));
                step += 1;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                var pos = ele.GetClickablePoint();
                Mouse.MoveTo(pos.X, pos.Y);
                Mouse.Click();

                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                step += 1;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                Thread.Sleep(1000);

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(0);
                step += 1;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();
                Thread.Sleep(3000);

                //var indexToReportsElement = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Index to Reports")));
                var indexToReportsElement = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByName("Daftar Laporan")));
                step += 1;
                if (indexToReportsElement == null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + indexToReportsElement.Properties.Name.ToString());
                indexToReportsElement.Focus();
                Thread.Sleep(2000);

                var reportMainTab = "";
                switch (reportType)
                {
                    case "sales":
                        reportMainTab = "Laporan Penjualan";
                        break;
                    case "ar":
                        reportMainTab = "Akun Piutang & Pelanggan";
                        break;
                    case "outlet":
                        reportMainTab = "Akun Piutang & Pelanggan";
                        break;
                    case "stock":
                        reportMainTab = "Persediaan";
                        break;
                    case "labarugi":
                        reportMainTab = "Laporan Keuangan";
                        break;
                    case "cashflow":
                        reportMainTab = "Kas & Bank";
                        break;
                    case "neraca":
                        reportMainTab = "Buku Besar";
                        break;
                }
                var reportElement1 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportMainTab));
                step += 1;
                if (reportElement1 == null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + reportElement1.Properties.Name.ToString());
                reportElement1.Click();
                Thread.Sleep(1000);

                var reportName = "";
                switch (reportType)
                {
                    case "sales":
                        reportName = "Rincian Penjualan per Pelanggan";
                        break;
                    case "ar":
                        reportName = "Rincian Pembayaran Faktur";
                        break;
                    case "outlet":
                        reportName = "Daftar Pelanggan";
                        break;
                    case "stock":
                        reportName = "Ringkasan Valuasi Persediaan";
                        break;
                    case "labarugi":
                        reportName = "Laba/Rugi (Standar)";
                        break;
                    case "cashflow":
                        reportName = "Arus Kas per Akun";
                        break;
                    case "neraca":
                        reportName = "Neraca Saldo";
                        break;
                }
                var reportElement2 = indexToReportsElement.FindFirstDescendant(cf.ByName(reportName));
                step += 1;
                if (reportElement2 == null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + reportElement2.Properties.Name.ToString());
                reportElement2.SetForeground();
                reportElement2.Focus();
                reportElement2.DoubleClick();
                Thread.Sleep(10000);

                // Opening Report Format Window
                AutomationElement reportFormatElement = null;
                auEle = window.FindAllChildren(cr => cr.ByName("Report Format"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        reportFormatElement = item;
                        break;
                    }
                }
                step += 1;
                if (reportFormatElement is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + reportFormatElement.Properties.Name.ToString());
                reportFormatElement.Focus();
                Thread.Sleep(2000);

                // Filters && Parameters => find it under 'reportFormatElement' windows tree
                var filtersAndParametersElement = reportFormatElement.FindFirstDescendant(cf.ByName("Filter && Parameter", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                step += 1;
                if (filtersAndParametersElement == null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + filtersAndParametersElement.Properties.Name.ToString());
                filtersAndParametersElement.Focus();
                Thread.Sleep(500);

                if (reportType == "labarugi")
                {
                    AutomationElement[] checkboxes = filtersAndParametersElement.FindAllDescendants(cr => cr.ByClassName("TCheckBox"));
                    foreach (AutomationElement checkbox in checkboxes)
                    {
                        switch (checkbox.Name)
                        {
                            case "Tampilkan Induk":
                                checkbox.AsCheckBox().IsChecked = true;
                                break;
                            case "Tampilkan Anak":
                                checkbox.AsCheckBox().IsChecked = false;
                                break;
                            case "Termasuk saldo nol":
                                checkbox.AsCheckBox().IsChecked = true;
                                break;
                            case "Tampilkan jumlah Induk":
                                checkbox.AsCheckBox().IsChecked = true;
                                break;
                            case "Tampilkan Total":
                                checkbox.AsCheckBox().IsChecked = false;
                                break;
                        }
                        Thread.Sleep(2000);
                    }
                }
                // Sending combo box value '<Semua>' when wwhen opening Stock Valuation //
                // "stock"
                if (reportType == "stock")
                {
                    var filterBarang = filtersAndParametersElement.FindFirstDescendant(cf.ByName("<Nihil>", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                    step += 1; 
                    if (filterBarang == null)
                    {
                        Log.Information($"[Step #{step}] Quitting, end of OpenReport function -> stock !!");
                        return false;
                    }
                    Log.Information("Element Interaction on property named -> " + filterBarang.Properties.Name.ToString());
                    //filterBarang.Focus();
                    filterBarang.AsTextBox().Text = "<Semua>\r";
                    Thread.Sleep(2000);
                }

                // TabDateFromTo
                var tabDateFromToElement = filtersAndParametersElement.FindFirstDescendant(cf.ByName("TabDate", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                step += 1; 
                if (tabDateFromToElement == null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of OpenReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + tabDateFromToElement.Properties.Name.ToString());
                tabDateFromToElement.Focus();

                // Sending Report Date Parameters //
                AutomationElement[] dateElements = tabDateFromToElement.FindAllDescendants(cf.ByClassName("TDateEdit"));

                if (dateElements.Length > 0)
                {
                    Log.Information($"Number of DATE parameters on screen is: {dateElements.Length}");

                    for (int index = dateElements.Length - 1; index > -1; index--)
                    {
                        var dateparam = "";
                        if (index != 0)
                        {
                            dateparam = DateManipultor.GetFirstDate() + "/" + DateManipultor.GetPrevMonth() + "/" + DateManipultor.GetPrevYear();
                            //SendingDate(dateElements[index], "01/01/2000");
                        }
                        else
                        {
                            dateparam = DateManipultor.GetLastDayOfPrevMonth() + "/" + DateManipultor.GetPrevMonth() + "/" + DateManipultor. GetPrevYear();
                            //SendingDate(dateElements[index], "31/12/2023");
                        }
                        SendingDate(dateElements[index], dateparam);

                    }
                }

                reportFormatElement.FindFirstDescendant(cf.ByName("OK")).AsButton().Click();
                return DownloadReport(reportType);
            }
            catch (Exception ex)
            {
                Log.Information(ex.ToString());
                return false;
            }
        }
        static bool IsFileExists(string path, string fileName)
        {
            string fullPath = Path.Combine(path, fileName);
            return File.Exists(fullPath);
        }

        private static bool SendingDate(AutomationElement ele, string date)
        {
            var step = 0;
            try
            {
                step = +1; 
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of SendingDate function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                ele.Click();

                // Send date parameter
                ele.AsTextBox().Enter("\b\b\b\b\b\b\b\b");
                ele.AsTextBox().Text = date;

                // TWinControl
                var childEle = ele.FindFirstDescendant(cf => cf.ByClassName("TWinControl"));
                step = +1;
                if (childEle is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of SendingDate function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + childEle.Properties.ClassName.ToString());
                childEle.Click();
                Thread.Sleep(500);
                childEle.Click();

                Log.Information($"Sending date parameter");

                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.Message);
                return false;
            }
        }

        static bool DownloadReport(string reportName)
        {
            Thread.Sleep(10000);
            var step = 0;
            try
            {
                /// Start downloading report process 
                /// Travesing back to accurate desktop main windows 
                AutomationElement ele1 = null;
                AutomationElement[] auEle = window.FindAllChildren(cf.ByName("ACCURATE 5", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele1 = item;
                        break;
                    }
                }
                step = +1;
                if (ele1 is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of DownloadReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                Thread.Sleep(500);

                var ele = ele1.FindFirstDescendant(cf => cf.ByName("PriviewToolBar"));
                step = +1; 
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of DownloadReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                //Export settings
                Thread.Sleep(1000);
                ele.FindFirstChild(cf.ByName("Expor", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring)).AsButton().Click();
                Thread.Sleep(1000);

                /// The export button action resulting new window opened 
                auEle = window.FindAllChildren(cf => cf.ByName("Export to Excel"));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        ele1 = item;
                        break;
                    }
                }
                step = +1;
                if (ele1 is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of DownloadReport function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                ele1.SetForeground();
                ele1.AsButton().Click();
                Thread.Sleep(2000);

                // Put here the code for iteration of report parameter check box //
                var exportToExcelElement = window.FindFirstChild(cf => cf.ByName("Export to Excel"));
                step = +1;
                if (exportToExcelElement is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of DownloadReport function !!");
                    return false;
                }
                AutomationElement[] checkboxes = exportToExcelElement.FindAllDescendants(cr => cr.ByClassName("TCheckBox"));
                foreach (AutomationElement checkbox in checkboxes)
                {
                    switch (checkbox.Name)
                    {
                        case "Pictures":
                            checkbox.AsCheckBox().IsChecked = false;
                            Thread.Sleep(2000);
                            break;
                        case "Merge cells":
                            checkbox.AsCheckBox().IsChecked = false;
                            Thread.Sleep(2000);
                            break;
                    }
                }
                // End of codes // 

                // Clicking OK button  //
                ele1.FindFirstChild(cf => cf.ByName("OK")).AsButton().Click();
                Log.Information("Clicking 'OK' button...");

                if (!SavingFileDialog(reportName))
                { return false; }

                return true;
            }
            catch (Exception ex) { Log.Information(ex.ToString()); return false; }
        }

        private static bool SavingFileDialog(string reportName)
        {
            var step = 0; 
            Thread.Sleep(5000);
            AutomationElement mainEle = null;
            AutomationElement[] auEle = window.FindAllChildren(cr => cr.ByName("Simpan Sebagai"));
            foreach (AutomationElement item in auEle)
            {
                if (item.Properties.ProcessId == pid)
                {
                    mainEle = item;
                    break;
                }
            }
            step = +1;
            if (mainEle is null)
            {
                Log.Information($"[Step #{step}] Quitting, end of SavingFileDialog function !!");
                return false;
            }
            Log.Information("Element Interaction on property named -> " + mainEle.Properties.Name.ToString());
            Thread.Sleep(500);

            AutomationElement ele1 = null;
            AutomationElement[] auEle1 = mainEle.FindAllDescendants(cr => cr.ByName("Nama File:", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
            foreach (AutomationElement item in auEle1)
            {
                if (item.Properties.ControlType.ToString() == "Edit")
                {
                    ele1 = item;
                    break;
                }
            }
            step = +1;
            if (ele1 is null)
            {
                Log.Information($"[Step #{step}] Quitting, end of SavingFileDialog function !!");
                return false;
            }
            Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
            Thread.Sleep(500);

            var excelname = "";
            switch (reportName)
            {
                case "sales":
                    excelname = "Sales_Data";
                    break;
                case "ar":
                    excelname = "Repayment_Data";
                    break;
                case "outlet":
                    excelname = "Master_Outlet";
                    break;
                case "stock":
                    excelname = "Laporan_Stock";
                    break;
                case "labarugi":
                    excelname = "Laporan_LabaRugi";
                    break;
                case "cashflow":
                    excelname = "Laporan_ArusKas";
                    break;
                case "neraca":
                    excelname = "Laporan_NeracaSaldo";
                    break;
            }
            ele1.SetForeground();
            ele1.Focus();
            ele1.AsTextBox().Enter($"{appfolder}\\{excelname}");
            Thread.Sleep(500);

            //Save
            AutomationElement ele2 = null;
            ele2 = mainEle.FindFirstChild(cf.ByName("Save"));
            step = +1;
            if (ele2 is null)
            {
                Log.Information($"[Step #{step}] Quitting, end of SavingFileDialog function !!");
                return false;
            }
            Log.Information("Element Interaction on property named -> " + ele2.Properties.Name.ToString());
            ele2.AsButton().Click();
            //* Pause the app to wait file saving is finnished //
            DateTime startTime = DateTime.Now;
            Thread.Sleep(1000);
            while (DateTime.Now - startTime < TimeSpan.FromMinutes(3))
            {
                if (IsFileExists(appfolder, excelname + ".xls"))
                {
                    //Log.Information("File saved successfully...");
                    break;
                }
                Thread.Sleep(5000);
            }
            if (!IsFileExists(appfolder, excelname + ".xls"))
            {
                Console.WriteLine("3 minutes timeout expired when saving file...");
                return false;
            }
            return true;
        }

        static bool ClosingWorkspace()
        {
            Thread.Sleep(2000);
            var step = 0;
            try
            {
                // Traversing back to accurate desktop main windows //
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllChildren(cf.ByName("ACCURATE 5", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                step++;
                if (mainElement is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of ClosingWorkspace function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(2000);

                var ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of ClosingWorkspace function !!");
                    return false;
                }
                Log.Information("Element Interaction on property with class named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                ele.Focus();
                Thread.Sleep(2000);

                //ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Windows")));
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Jendela")));
                step++; 
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of ClosingWorkspace function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                var pos = ele.GetClickablePoint();
                Mouse.MoveTo(pos.X, pos.Y);
                Mouse.Click();
                Thread.Sleep(2000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                step++; 
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of ClosingWorkspace function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());

                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(1);
                step++; 
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of ClosingWorkspace function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.Click();

                Thread.Sleep(1000);

                return true;
            }
            catch (Exception ex)
            {
                Log.Information(ex.ToString());
                return false;
            }
        }

        private static bool CloseApp()
        {
            Thread.Sleep (2000);
            var step = 0;
            try
            {
                AutomationElement mainElement = null;
                AutomationElement[] auEle = window.FindAllChildren(cf.ByName("ACCURATE 5", FlaUI.Core.Definitions.PropertyConditionFlags.MatchSubstring));
                foreach (AutomationElement item in auEle)
                {
                    if (item.Properties.ProcessId == pid)
                    {
                        mainElement = item;
                        break;
                    }
                }
                step++;
                if (mainElement is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + mainElement.Properties.Name.ToString());
                mainElement.SetForeground();
                Thread.Sleep(500);

                var ele = WaitForElement(() => mainElement.FindFirstDescendant(cr => cr.ByClassName("TsuiSkinMenuBar")));
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.ClassName.ToString());
                //ele.SetForeground();
                Thread.Sleep(500);

                /* Click on Reports menu */
                ele = WaitForElement(() => ele.FindFirstDescendant(cr => cr.ByName("Berkas")));
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                var pos = ele.GetClickablePoint();
                Mouse.MoveTo(pos.X, pos.Y);
                Mouse.Click();
                Thread.Sleep(1000);

                // Context
                ele = WaitForElement(() => window.FindFirstDescendant(cr => cr.ByName("Context")));
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele.Properties.Name.ToString());
                Thread.Sleep(1000);

                // Exit - MenuItem #13
                ele = ele.FindAllDescendants((cr => cr.ByControlType(FlaUI.Core.Definitions.ControlType.MenuItem))).ElementAt(12);
                step++;
                if (ele is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named 'Context' with Child id# -> " + ele.Properties.AutomationId.ToString());
                ele.AsMenuItem().Focus();
                Thread.Sleep(1000);
                ele.Click();
                Thread.Sleep(1000);

                /* The Menu 'File' -> 'Close' clicked action resulting new window opened */
                var ele1 = window.FindFirstDescendant(cf => cf.ByName("Confirm"));
                step++;
                if (ele1 is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());
                Thread.Sleep(1000);

                /* Clicking 'Yes' button  */
                ele1.FindFirstChild(cf => cf.ByName("Yes")).AsButton().Click();
                Log.Information("Clicking 'Yes' button...");
                Thread.Sleep(5000);

                /* The Menu 'Yes' button clicked action resulting new window opened */
                ele1 = window.FindFirstDescendant(cf => cf.ByName("Confirm"));
                step++;
                if (ele1 is null)
                {
                    Log.Information($"[Step #{step}] Quitting, end of CloseApp function !!");
                    return false;
                }
                Log.Information("Element Interaction on property named -> " + ele1.Properties.Name.ToString());

                /* Clicking OK button  */
                ele1.FindFirstChild(cf => cf.ByName("No")).AsButton().Click();
                Log.Information("Clicking 'No' button...");
                Thread.Sleep(1000);
                return true;
            }
            catch (Exception ex)
            {
                Log.Information($"Exception: {ex.ToString()}");
                return false;
            }
        }

        static bool ZipandSend()
        {
            try
            {
                Log.Information("Starting zipping file reports process...");
                var strDsPeriod = DateManipultor.GetPrevYear() + DateManipultor.GetPrevMonth();

                Log.Information("Moving standart excel reports file to uploaded folder...");
                // move excels files to Datafolder
                var path = appfolder + @"\Master_Outlet.xls";
                var path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_OUTLET.xls";
                File.Move(path, path2, true);
                path = appfolder + @"\Sales_Data.xls";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_SALES.xls";
                File.Move(path, path2, true);
                path = appfolder + @"\Repayment_Data.xls";
                path2 = uploadfolder + @"\ds-" + dtID + "-" + dtName + "-" + strDsPeriod + "_AR.xls";
                File.Move(path, path2, true);

                // set zipping name for files
                Log.Information("Zipping Transaction file(s)");
                var strZipFile = dtID + "-" + dtName + "_" + strDsPeriod + ".zip";
                ZipFile.CreateFromDirectory(uploadfolder, sharingfolder + Path.DirectorySeparatorChar + strZipFile);

                // Send the ZIP file to the API server 
                Log.Information("Sending ZIP file to the API server...");
                var strStatusCode = "0"; // variable for debugging cUrl test
                using (cUrlClass myCurlCls = new cUrlClass('Y', issandbox.ToArray().First(), "", sharingfolder + Path.DirectorySeparatorChar + strZipFile))
                {
                    strStatusCode = myCurlCls.SendRequest();
                    if (strStatusCode == "200")
                    {
                        Log.Information("DATA TRANSACTION SHARING - SELESAI");
                    }
                    else
                    {
                        Log.Information("Failed to send TRANSACTION file to API server... => " + strStatusCode);
                    }
                }

                /* Ending logging before sending log file to API server */
                Log.CloseAndFlush();
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Sending log file to the API server..."));
                strStatusCode = "0"; // variable for debugging cUrl test
                using (cUrlClass myCurlCls = new cUrlClass('Y', issandbox.ToArray().First(), "", appfolder + Path.DirectorySeparatorChar + logfilename))
                {
                    strStatusCode = myCurlCls.SendRequest();
                    if (strStatusCode != "200")
                    {
                        throw new Exception($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Failed to send LOG file to API server...");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] Error during ZIP and cUrl send => {ex.Message}"));
                return false;
            }
        }
    }
}