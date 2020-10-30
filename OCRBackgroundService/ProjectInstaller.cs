using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;

namespace OCRBackgroundService
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : System.Configuration.Install.Installer
    {
        public ProjectInstaller()
        {           
            InitializeComponent();
            var config = ConfigurationManager.OpenExeConfiguration(this.GetType().Assembly.Location);
            this.serviceInstaller1.Description = config.AppSettings.Settings["SERVICEDESC"].Value;
            this.serviceInstaller1.ServiceName = config.AppSettings.Settings["SERVICENAME"].Value;
        }
    }
}
