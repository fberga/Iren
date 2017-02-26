using System.Configuration;

namespace Iren.PSO.Core
{
    public class CryptHelper
    {
        /// <summary>
        /// Funzione per criptare le sezioni del file di configurazione che contengono dati sensibili.
        /// </summary>
        /// <param name="sections">Lista di sezioni da criptare.</param>
        public static void CryptSection(params string[] sections)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            string provider = "RsaProtectedConfigurationProvider";

            foreach (string sectionName in sections)
            {
                ConfigurationSection section = config.GetSection(sectionName);
                if (section != null)
                {
                    if (!section.SectionInformation.IsProtected)
                    {
                        if (!section.ElementInformation.IsLocked)
                        {
                            section.SectionInformation.ProtectSection(provider);

                            section.SectionInformation.ForceSave = true;
                            config.Save(ConfigurationSaveMode.Modified);
                        }
                    }
                }
            }
        }
    }
}
