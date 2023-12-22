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
        static string LoginId = ConfigurationManager.AppSettings["loginId"];
        static string LoginPassword = ConfigurationManager.AppSettings["password"];
        static string appExe = ConfigurationManager.AppSettings["erpappnamepath"];
        static string dbserveraddr = ConfigurationManager.AppSettings["dbserveraddress"].ToUpper();
        static string issandbox = ConfigurationManager.AppSettings["uploadtosandbox"].ToUpper();
        static string enableconsolelog = ConfigurationManager.AppSettings["enableconsolelog"].ToUpper();
        static string appfolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\" + ConfigurationManager.AppSettings["appfolder"];
        static string uploadfolder = appfolder + @"\" + ConfigurationManager.AppSettings["uploadfolder"];
        static string sharingfolder = appfolder + @"\" + ConfigurationManager.AppSettings["sharingfolder"];
        static string iposgudang = ConfigurationManager.AppSettings["namagudangdiipos"].ToUpper();
        static string shortcuttoipos = ConfigurationManager.AppSettings["shortcuttoipos"].ToUpper();
        static string dbname = ConfigurationManager.AppSettings["dbname"];
        //static string screenshotfolder = appfolder + @"\" + ConfigurationManager.AppSettings["screenshotfolder"];
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
                Console.SetWindowPosition(0, 0);
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

                /*
                if (!OpenAppAndDBConfig())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (OpenAppAndDBConfig) !!!");
                    return;
                }
                if (!LoginApp())
                {
                    Console.Beep();
                    Task.Delay(500);
                    Log.Information("application automation failed when running app (LoginApp) !!!");
                    return;
                }
                Console.WriteLine("Hello World!");
                */
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

                Task.Run(() => Console.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")} INF] iPOS ver.4 Automation - ***   END   ***"));
                if (automationUIA3 is not null)
                {
                    automationUIA3.Dispose();
                }
                Log.CloseAndFlush();
            }
        }
    }
}