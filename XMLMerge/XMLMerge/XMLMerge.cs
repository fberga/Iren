using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace XMLMerge
{
    public partial class XMLMerge : Form
    {
        string _path = "";

        Dictionary<string, Dictionary<DateTime, List<Tuple<XElement, XElement>>>> _tot = new Dictionary<string, Dictionary<DateTime, List<Tuple<XElement, XElement>>>>();

        public XMLMerge()
        {
            InitializeComponent();
        }

        private void XMLMerge_Shown(object sender, EventArgs e)
        {
            
        
        }

        private void Parse()
        {
            string[] list = Directory.GetFiles(_path);
            //ottengo un albero con tutti i valori indicizzati per UP e per giorno
            txtOutput.AppendText("Trovat" + (list.Length == 1 ? "o " : "i ") + list.Length + " file \r\n");
            foreach (string file in list)
            {
                XMLParser p = new XMLParser();
                txtOutput.AppendText("Parsing '" + file + "' ");
                if (p.Parse(file))
                {
                    DateTime d = new DateTime(p.Anno, p.Mese, p.Giorno);
                    if (!_tot.ContainsKey(p.UnitaProduzione))
                        _tot.Add(p.UnitaProduzione, new Dictionary<DateTime, List<Tuple<XElement, XElement>>>());

                    if (!_tot[p.UnitaProduzione].ContainsKey(d))
                        _tot[p.UnitaProduzione].Add(d, new List<Tuple<XElement, XElement>>());

                    _tot[p.UnitaProduzione][d].Add(Tuple.Create(p.GiornoContrPositivo, p.GiornoContrNegativo));

                    txtOutput.AppendText("fatto!\r\n");
                }
                else
                {
                    txtOutput.AppendText("fallito...\r\n");
                }
            }
        }

        private void PrepareOutput()
        {
            foreach (var up in _tot)
            {
                Dictionary<string, XDocument> o = new Dictionary<string, XDocument>();
                txtOutput.AppendText("Scrivendo output per " + up.Key);

                XDocument header = new XDocument(
                    new XDeclaration("1.0", "UTF-8", ""),
                    new XElement("Periodo", new XAttribute("AnnoRif", ""), new XAttribute("MeseRif", ""),
                        new XElement("UP", new XAttribute("CODICE", up.Key),
                            new XElement("ContributoPositivo"),
                            new XElement("ContributoNegativo"))));

                foreach (var giorno in up.Value)
                {
                    if (!o.ContainsKey(giorno.Key.ToString("yyyy-MM")))
                    {
                        XDocument tmp = new XDocument(header);
                        tmp.Element("Periodo").Attribute("AnnoRif").Value = giorno.Key.ToString("yyyy");
                        tmp.Element("Periodo").Attribute("MeseRif").Value = giorno.Key.ToString("MM");
                        o.Add(giorno.Key.ToString("yyyy-MM"), tmp);
                    }

                    foreach (var contr in giorno.Value)
                    {
                        (o[giorno.Key.ToString("yyyy-MM")].Descendants("ContributoPositivo").First() as XElement).Add(contr.Item1);
                        (o[giorno.Key.ToString("yyyy-MM")].Descendants("ContributoNegativo").First() as XElement).Add(contr.Item2);
                    }
                }

                foreach (var doc in o)
                {
                    string filename = doc.Key + (_tot.Count > 1 ? "_" + up.Key : "") + ".xml";
                    doc.Value.Save(Path.Combine(_path, filename));
                }
                txtOutput.AppendText(" fatto!\r\n");
            }
        }

        private void btnApri_Click(object sender, EventArgs e)
        {
            if (scegliCartella.ShowDialog() == DialogResult.OK)
            {
                _path = scegliCartella.SelectedPath;
                txtPercorso.Text = _path;
                Parse();
            }
        }

        private void btnEsegui_Click(object sender, EventArgs e)
        {
            PrepareOutput();
        }
    }
}
