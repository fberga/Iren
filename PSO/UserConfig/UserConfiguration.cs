using System.Configuration;

namespace Iren.PSO.UserConfig
{
    public class UserConfiguration : ConfigurationSection
    {
        [ConfigurationProperty("", IsDefaultCollection = true)]
        public UserConfigCollection Items
        {
            get { return (UserConfigCollection)base[""]; }
            set { base[""] = value; }
        }
    }
}
