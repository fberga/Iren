using System;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreControllo : Form
    {
        public string CtrlName { get { return txtName.Text; } }
        public string CtrlText { get { return txtLabel.Text; } }

        public ConfiguratoreControllo(Control ribbon, Type t)
        {
            InitializeComponent();

            string prefix = "";

            if (t == typeof(RibbonDropDown))
                prefix = RibbonDropDown.NEW_COMBO_PREFIX;
            else if (t == typeof(RibbonGroup))
                prefix = RibbonGroup.NEW_GROUP_PREFIX;

            int prog = Utility.FindLastOfItsKind(ribbon, prefix, t) + 1;
            txtLabel.Text = prefix + " " + prog;
            txtName.Text = txtLabel.Text.Replace(" ", "_");
        }
        public ConfiguratoreControllo(Control ctrl, Control ribbon, Type t)
        {
            InitializeComponent();

            txtLabel.Text = ctrl.Text;
            txtName.Text = ctrl.Name;
        }

        private void Applica_Click(object sender, EventArgs e)
        {
            if (txtLabel.Text == "" || txtName.Text == "")
            {
                MessageBox.Show("Inserire un nome e/o un label per il controllo.", "ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }

        private void Annulla_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            Close();
        }
    }
}
