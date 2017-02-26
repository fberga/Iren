using Iren.PSO.Base;
using Iren.PSO.UserConfig;
using System;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace Iren.PSO.Forms
{
    public partial class FormConfiguraPercorsi : Form
    {
        DataTable _dt = new DataTable("usrConfig") 
        {
            Columns =
            {
                {"Key", typeof(string)},
                {"Proprietà", typeof(string)},
                {"Produzione", typeof(string)},
                {"Test", typeof(string)},
                {"Emergenza", typeof(string)}
            }
        };

        public FormConfiguraPercorsi()
        {
            InitializeComponent();

            dataGridConfigurazioni.DataSource = _dt;
            
            this.Text = Simboli.NomeApplicazione + " - Configura percorsi";

            int width = dataGridConfigurazioni.Width * 93 / 100;

            dataGridConfigurazioni.Columns[0].Visible = false;

            dataGridConfigurazioni.Columns[1].Width = (width / 4);
            dataGridConfigurazioni.Columns[1].ReadOnly = true;
            dataGridConfigurazioni.Columns[2].ReadOnly = true;
            dataGridConfigurazioni.Columns[1].DefaultCellStyle = new DataGridViewCellStyle() 
            {
                SelectionBackColor = System.Drawing.Color.White,
                SelectionForeColor = System.Drawing.Color.Black,
                Font = new Font(dataGridConfigurazioni.Font, FontStyle.Bold)
            };
            
            dataGridConfigurazioni.Columns[2].Width = (width / 4);
            
            dataGridConfigurazioni.Columns[3].Width = (width / 4);
            dataGridConfigurazioni.Columns[4].Width = (width / 4);
            //dataGridConfigurazioni.Columns[3].ReadOnly = true;
            
        }

        private void FormConfig_Load(object sender, EventArgs e)
        {
            var config = (UserConfiguration)ConfigurationManager.GetSection("usrConfig");

            foreach (UserConfigElement item in config.Items)
            {
                if (item.Visibile)
                    _dt.Rows.Add(item.Key, item.Desc, item.Value, item.Test, item.Emergenza);
            }

        }

        private void btnApplica_Click(object sender, EventArgs e)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var section = (UserConfiguration)config.GetSection("usrConfig");

            foreach (DataRow r in _dt.GetChanges().Rows)
            {
                section.Items[r["Key"].ToString()].Value = r["Produzione"].ToString();
                section.Items[r["Key"].ToString()].Test = r["Test"].ToString();
                section.Items[r["Key"].ToString()].Emergenza = r["Emergenza"].ToString();
            }

            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("usrConfig");
        }
    }
}
