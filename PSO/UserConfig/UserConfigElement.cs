using System;
using System.Configuration;

namespace Iren.PSO.UserConfig
{
    public class UserConfigElement : ConfigurationElement
    {
        public enum ElementType
        {
            path, pathNoCheck, to, subject, body, nameFormat
        }


        [ConfigurationProperty("key", IsRequired = true, IsKey = true)]
        public string Key
        {
            get { return (string)base["key"]; }
            set { base["key"] = value; }
        }

        [ConfigurationProperty("type", IsRequired = true, IsKey = true)]
        public ElementType Type
        {
            get { return (ElementType)Enum.Parse(typeof(ElementType), base["type"].ToString()); }
            set { base["type"] = value; }
        }

        [ConfigurationProperty("desc", IsRequired = false, DefaultValue="")]
        public string Desc
        {
            get { return (string)base["desc"]; }
            set { base["desc"] = value; }
        }

        [ConfigurationProperty("value", IsRequired = true)]
        public string Value
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["value"]); }
            set { base["value"] = value; }
        }

        //[ConfigurationProperty("default", IsRequired = false, DefaultValue="")]
        //public string Default
        //{
        //    get { return Environment.ExpandEnvironmentVariables((string)base["default"]); }
        //    set { base["default"] = value; }
        //}

        [ConfigurationProperty("emergenza", IsRequired = false, DefaultValue="")]
        public string Emergenza
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["emergenza"]); }
            set { base["emergenza"] = value; }
        }

        [ConfigurationProperty("archivio", IsRequired = false, DefaultValue = "")]
        public string Archivio
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["archivio"]); }
            set { base["archivio"] = value; }
        }

        [ConfigurationProperty("test", IsRequired = false, DefaultValue = "")]
        public string Test
        {
            get { return Environment.ExpandEnvironmentVariables((string)base["test"]); }
            set { base["test"] = value; }
        }

        [ConfigurationProperty("visibile", IsRequired = false, DefaultValue="true")]
        public bool Visibile
        {
            get { return (bool)base["visibile"]; }
            set { base["visibile"] = value; }
        }
    }
}
