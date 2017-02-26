using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;

namespace Iren.PSO.Services.PollingService
{
    public partial class PollingService : ServiceBase
    {
        FileSystemWatcher _fsw = new FileSystemWatcher();
        string _path = "";

        public PollingService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            _path = ConfigurationManager.AppSettings["pollingpath"];

            _fsw.Path = _path;
            _fsw.NotifyFilter = NotifyFilters.LastWrite;
            _fsw.Changed += ChangeInWatchedFile;
        }

        private void ChangeInWatchedFile(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType == WatcherChangeTypes.Changed)
            {
                System.Windows.Forms.MessageBox.Show("FileModificato");
            }
        }

        protected override void OnStop()
        {
            _fsw.Changed -= ChangeInWatchedFile;
        }
    }
}
