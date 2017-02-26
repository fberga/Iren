using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace XMLMerge
{
    class XMLParser
    {
        XDocument _doc;
        public XElement GiornoContrPositivo { get; private set; }
        public XElement GiornoContrNegativo { get; private set; }

        public string UnitaProduzione { get; private set; }
        public int Anno { get; private set; }
        public int Mese { get; private set; }
        public int Giorno { get; private set; }
        
        public bool Parse(string path)
        {
            _doc = XDocument.Load(path);

            if(Validate())
                return true;


            return false;
        }

        private bool Validate()
        {
            if(_doc.Nodes().Count() > 1)
                return false;

            XElement periodo = _doc.FirstNode as XElement;

            try
            {
                Anno = int.Parse(periodo.Attribute("AnnoRif").Value);
                Mese = int.Parse(periodo.Attribute("MeseRif").Value);
            }
            catch
            {
                return false;
            }

            if (periodo.Nodes().Count() > 1)
                return false;

            XElement up = periodo.FirstNode as XElement;

            try
            {
                UnitaProduzione = up.Attribute("CODICE").Value;
            }
            catch
            {
                return false;
            }

            if (up.Elements().Count() != 2)
                return false;

            try
            {
                if (up.Elements("ContributoPositivo").Count() > 1)
                    return false;
            }
            catch
            {
                return false;
            }
            try
            {
                if (up.Elements("ContributoNegativo").Count() > 1)
                    return false;
            }
            catch
            {
                return false;
            }

            try
            {
                if (up.Elements("ContributoNegativo").Elements("Giorno").Count() > 1)
                    return false;
            }
            catch
            {
                return false;
            }
            
            try
            {
                if (up.Elements("ContributoNegativo").Elements("Giorno").Count() > 1)
                    return false;
            }
            catch
            {
                return false;
            }

            if (up.Elements("ContributoPositivo").Descendants().Count() != 1)
                return false;

            GiornoContrPositivo = up.Elements("ContributoPositivo").Elements("Giorno").First();
            
            if (up.Elements("ContributoNegativo").Descendants().Count() != 1)
                return false;
            
            GiornoContrNegativo = up.Elements("ContributoNegativo").Elements("Giorno").First();

            try
            {
                Giorno = int.Parse(GiornoContrPositivo.Attribute("ID").Value);
            }
            catch
            {
                return false;
            }

            return true;
        }
    }
}
