using Microsoft.Owin.Hosting;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace OwinSelfhost
{
    public class Program
    {
        [DllImport("Kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();
        [DllImport("User32.dll")]
        private static extern IntPtr ShowWindow(IntPtr hWnd, int cmdShow);
        public static System.Timers.Timer _timer;

        static void Main()
        {
            DeleteOldLogFiles();
            WriteToLog(string.Format("OwinSelfHost version {0} Started ***", GetAssemblyFileVersion()));
            IntPtr hWnd = GetConsoleWindow();
            // hide console window and taskbar
            ShowWindow(hWnd, 0);

            using (Mutex mutex = new Mutex(false, "Global\\" + appGuid))
            {
                if (!mutex.WaitOne(0, false))
                {
                    return;
                }                
            }

            int port = 9000;
            StartHost(ref port);
        }

        /// <summary>
        /// gets the assembly version
        /// </summary>
        /// <returns></returns>
        public static string GetAssemblyFileVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fileVersion = FileVersionInfo.GetVersionInfo(assembly.Location);
            return fileVersion.FileVersion;
        }

        /// <summary>
        /// sets 10 minute interval timer
        /// </summary>
        /// <param name="port"></param>
        static void SetTimer(int port)
        {
            _timer = new System.Timers.Timer();
            _timer.Interval = GetConfigValue("PortTimerInterval", 600000);
            _timer.Elapsed += (sender, e) => { HandleTimerElapsed(port); };
            _timer.Enabled = true;
        }

        static void HandleTimerElapsed(int port)
        {
            SetPortCookie(port);
        }

        /// <summary>
        /// loops domains and sets cookie for each one
        /// </summary>
        /// <param name="port"></param>
        private static void SetPortCookie(int port)
        {
            try
            {
                string[] domains = ConfigurationManager.AppSettings["Domain"].Split(';');

                foreach(string domain in domains)
                {
                    SetPortCookieForDomain(domain, port);
                }
            }
            catch (Exception ex)
            {
                Program.WriteErrorToLog(ex);
            }
        }

        private static void SetPortCookieForDomain(string domain, int port)
        {
            try
            {
                WriteToLog(string.Format("SetPortCookie with port {0} for domain {1}", port, domain));
                SHDocVw.InternetExplorer ie = new SHDocVw.InternetExplorer();
                object url = string.Format("http://{0}/FlyingCargo/Listener.aspx?port={1}", domain, port);
                ie.Visible = false;
                ie.Silent = true;
                ie.Navigate2(ref url);
                Thread.Sleep(GetConfigValue("WaitForIeCall", 120000));
                ie.Quit();
                WriteToLog("SetPortCookie finished (Killed ie)");
            }
            catch (Exception ex)
            {
                Program.WriteErrorToLog(ex);
            }
        }

        /// <summary>
        /// Start OWIN host
        /// if port is taken, try the nxt one.
        /// </summary>
        /// <param name="port"></param>
        private static void StartHost(ref int port)
        {
            string baseAddress = string.Format("http://localhost:{0}/", port);

            try
            {
                using (WebApp.Start<Startup>(url: baseAddress))
                {
                    WriteToLog(string.Format("Registered port {0}", port));
                    // Create HttpCient and make a request to api/values 
                    HttpClient client = new HttpClient();

                    var response = client.GetAsync(baseAddress + "api/values").Result;
                    SetPortCookie(port);

                    if (Program.GetConfigValue("IsUseTimer", 1) == 1)
                    {
                        SetTimer(port);
                    }
                    
                    Console.WriteLine(response);
                    Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                    // keep listener alive
                    Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                string portTakenMessage = string.Format("Failed to listen on prefix '{0}' because it conflicts with an existing registration on the machine.", baseAddress);
                

                if (ex.InnerException != null && ex.InnerException.Message == portTakenMessage)
                {
                    Program.WriteToLog(string.Format("StartHost Error: {0}", ex.InnerException.Message));
                    // another instance of this listener is already running on this machine using this port. Try the next port. 
                    port += 1;
                    StartHost(ref port);
                }
                else
                {
                    Program.WriteErrorToLog(ex);
                }
            }
        }

        static public int GetConfigValue(string name, int defaultValue)
        {
            if (ConfigurationManager.AppSettings[name] == null)
            {
                return defaultValue;
            }
            else
            {
                return Convert.ToInt32(ConfigurationManager.AppSettings[name].ToString());
            }
        }

        static public void WriteErrorToLog(Exception ex)
        {
            var methodname = GetCallForExceptionThisMethod(MethodBase.GetCurrentMethod(), ex);

            WriteToLog(string.Format("{0} Error: {1}{2} *** StackTrace: {3}", methodname, ex.Message, Environment.NewLine, ex.StackTrace));
        }

        private static string GetCallForExceptionThisMethod(MethodBase methodBase, Exception e)
        {
            StackTrace trace = new StackTrace(e);
            StackFrame previousFrame = null;

            foreach (StackFrame frame in trace.GetFrames())
            {
                if (frame.GetMethod() == methodBase)
                {
                    break;
                }

                previousFrame = frame;
            }

            return previousFrame != null ? previousFrame.GetMethod().Name : null;
        }

        static public void WriteToLog(string message)
        {
            var logMessage = new StringBuilder();

            logMessage.Append(Environment.NewLine);
            logMessage.AppendFormat("{0} *** {1}", DateTime.Now.ToString(), message);            
            File.AppendAllText(GetPath(), logMessage.ToString());
        }

        /// <summary>
        /// creates OwinSelfHost folder if does not exist
        /// </summary>
        /// <returns></returns>
        static private string GetDirectory()
        {
            var path = Environment.ExpandEnvironmentVariables("%LOCALAPPDATA%") + "\\OwinSelfHost\\";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            return path;
        }

        /// <summary>
        /// returns c:\Users\username\AppData\Local\OwinSelfHost\OwinSelfHost.log
        /// </summary>
        /// <returns></returns>
        static private string GetPath()
        {           
            var dir = GetDirectory();
            var logFileName = string.Format("OwinSelfHost_{0}.log", DateTime.Now.ToString("dd.MM.yyyy"));

            return string.Concat(dir, logFileName);
        }

        /// <summary>
        /// deletes files last accessed before value in app config (LogDays).
        /// </summary>
        static void DeleteOldLogFiles()
        {
            try
            {
                var dir = GetDirectory();
                string[] files = Directory.GetFiles(dir);

                foreach (string file in files)
                {
                    FileInfo fi = new FileInfo(file);
                    if (fi.LastAccessTime < DateTime.Now.AddDays(-GetConfigValue("LogDays", 7)))
                    {
                        fi.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                Program.WriteErrorToLog(ex);
            }
        }

        private static string appGuid = "c0a76b5a-12ab-45c5-b9d9-d693faa6e7b9";
    }
}