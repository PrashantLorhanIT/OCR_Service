using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using log4net;
using log4net.Config;

namespace OCRBackgroundService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
               
        public Service1()
        {           
            InitializeComponent();
            var config = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);
            if (config != null)
            {
                if (config.AppSettings.Settings["SERVICENAME"] != null)
                {
                    this.ServiceName = config.AppSettings.Settings["SERVICENAME"].Value;
                }
                if (config.AppSettings.Settings["INTERVAL"] != null)
                {
                    this.timer.Interval = Convert.ToInt32(!string.IsNullOrEmpty(config.AppSettings.Settings["INTERVAL"].Value) ? config.AppSettings.Settings["INTERVAL"].Value : "10000");
                }
            }
        }

        protected override void OnStart(string[] args)
        {
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);          
            timer.Enabled = true;         
            WriteToFile("Service is started at " + DateTime.Now);
        }

        private void OnElapsedTime(object sender, ElapsedEventArgs e)
        {            
            if (timer.Enabled)
            {
                try
                {                  
                    WriteToFile("Service Elapsed Time" + DateTime.Now);
                    OCRUtility ocrUtility = new OCRUtility();
                    ocrUtility.DirectorySearch();
                    this.timer.Enabled = false;                   
                }
                catch (Exception ex)
                {                   
                  WriteToFile("Stack Trac:" + ex.StackTrace + "Message:" + ex.Message);
                }
                finally
                {
                    this.timer.Enabled = true;
                }
            }
        }

        protected override void OnStop()
        {
            timer.Stop();           
            WriteToFile("Service is stopped at " + DateTime.Now);           
        }

        public static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}
