using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Iren.ToolsExcel.ManifestEnhancer
{
    class ManifestEnhancer
    {
        static void Main(string[] args)
        {
            ManifestEnhancer mEnh = new ManifestEnhancer();
            mEnh.Run();
        }


        public void Run()
        {
            string basePath = @"Y:\PSO\";
            string repositoryPath = @"D:\Repository\Iren\ToolsExcel\";
            string[] projectFolders = Directory.GetDirectories(basePath);

            //carico la versione della libreria
            FileCopy.ToolsExcelFileCopy f = new FileCopy.ToolsExcelFileCopy();
            Version v = f.Version;

            foreach (string pF in projectFolders)
            {                
                if (pF.Contains("Sistema Comandi"))
                {
                    string vstoFile = Directory.GetFiles(pF, "*.vsto")[0];
                    string assemblyName = vstoFile.Split('\\').Last().Split('.').First();
                    string temporaryKey = Directory.GetFiles(Path.Combine(repositoryPath, assemblyName), "*TemporaryKey.pfx")[0];
                    string lastVersionFolder = Directory.GetDirectories(Path.Combine(pF, "Application Files")).Last();
                    string manifestFile = Directory.GetFiles(lastVersionFolder, "*.dll.manifest")[0];

                    string key = Path.Combine(lastVersionFolder, temporaryKey.Split('\\').Last());
                    File.Copy(temporaryKey, key, true);

                    string fileContents = System.IO.File.ReadAllText(manifestFile);
                    if (!fileContents.Contains("ToolsExcelFileCopy"))
                        //aggiorno il file
                        fileContents = fileContents.Replace("<vstav3:update enabled=\"true\" />", "<vstav3:update enabled=\"true\" />\n<vstav3:postActions>\n<vstav3:postAction>\n<vstav3:entryPoint class=\"Iren.ToolsExcel.FileCopy.ToolsExcelFileCopy\">\n<assemblyIdentity name=\"ToolsExcelFileCopy\" version=\"" + v.ToString() + "\" language=\"neutral\" processorArchitecture=\"msil\" />\n</vstav3:entryPoint>\n<vstav3:postActionData />\n</vstav3:postAction>\n</vstav3:postActions>");

                    System.IO.File.WriteAllText(manifestFile, fileContents);
                        
                    //creazione stringhe da eseguire
                    string bat = "mage -sign \"" + manifestFile + "\" -certfile \"" + key + "\"\n" +
                        "mage -update \"" + vstoFile + "\" -appmanifest \"" + manifestFile + "\" -certfile \"" + key +"\"\n";

                    var p = Process.Start(@"C:\Program Files (x86)\Microsoft Visual Studio 12.0\Common7\Tools\VsDevCmd.bat");
                    p.Close();
                }   
            }
        }

    }
}
