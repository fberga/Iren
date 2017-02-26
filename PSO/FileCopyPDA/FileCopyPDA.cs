using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Deployment;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Iren.PSO.PostDeployment
{
    public class FileCopyPDA : IAddInPostDeploymentAction
    {
        public void Execute(AddInPostDeploymentActionArgs args)
        {   
            XElement parameters = XElement.Parse(args.PostActionManifestXml);

            //configurabili
            string dataDirectory = @"Data\";
            string file = parameters.Attribute("filename").Value;

            //statici
            string sourcePath = args.AddInPath;
            string destPath = Environment.ExpandEnvironmentVariables(Base.Simboli.LocalBasePath);
            Uri deploymentManifestUri = args.ManifestLocation;
            string sourceFile = Path.Combine(sourcePath, dataDirectory, file);
            string destFile = Path.Combine(destPath, file);

            switch (args.InstallationStatus)
            {
                case AddInInstallationStatus.InitialInstall:
                    if (!Directory.Exists(destPath))
                        Directory.CreateDirectory(destPath);

                    System.IO.File.Copy(sourceFile, destFile, true);

                    if(ServerDocument.IsCustomized(destFile))
                        ServerDocument.RemoveCustomization(destFile);

                    ServerDocument.AddCustomization(destFile, deploymentManifestUri);

                    break;
                case AddInInstallationStatus.Update:
                    string dirUPDATE = Path.Combine(destPath, "UPDATE");
                    string fileUPDATE = Path.Combine(dirUPDATE, file);
                    if (!Directory.Exists(dirUPDATE))
                        Directory.CreateDirectory(dirUPDATE);

                    System.IO.File.Copy(sourceFile, fileUPDATE, true);

                    if (ServerDocument.IsCustomized(fileUPDATE))
                        ServerDocument.RemoveCustomization(fileUPDATE);

                    ServerDocument.AddCustomization(fileUPDATE, deploymentManifestUri);

                    break;
                case AddInInstallationStatus.Uninstall:
                    if (System.IO.File.Exists(destFile))
                    {
                        //rimuovo file di installazione
                        System.IO.File.Delete(destFile);

                        //rimuovo directory di update
                        string update = Path.Combine(destPath, "UPDATE");
                        if (Directory.Exists(update) && !Directory.EnumerateFileSystemEntries(update).Any())
                            Directory.Delete(update);

                        //rimuovo directory PSO
                        if (!Directory.EnumerateFileSystemEntries(destPath).Any())
                            Directory.Delete(destPath);
                    }
                    break;
            }
        }
    }
}