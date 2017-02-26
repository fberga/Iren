using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace Iren.PSO.Launcher
{
    static class Program
    {
        /// <summary>
        /// Punto di ingresso principale dell'applicazione.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //controllo se ci sono altre istanze del processo attive e le spengo
            var processes = Process.GetProcesses()
                .Where(p => p.ProcessName == Process.GetCurrentProcess().ProcessName && p.Id != Process.GetCurrentProcess().Id);

            foreach(Process p in processes)
            {
                p.Close();
            }            

            //imposto gli handler per l'update
            //ApplicationDeployment.CurrentDeployment.CheckForUpdateCompleted += Launcher.CurrentDeployment_CheckForUpdateCompleted;
            //ApplicationDeployment.CurrentDeployment.UpdateCompleted += Launcher.CurrentDeployment_UpdateCompleted;

            //Launcher.CheckForUpdates();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new LDaemon());
        }






    }
}
