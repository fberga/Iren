using System;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreTasto : Form
    {
        RibbonButton _btn;

        public ConfiguratoreTasto(RibbonButton btn, Control ribbon)
            : this(btn)
        {
            int prog = Utility.FindLastOfItsKind(ribbon, RibbonButton.NEW_BUTTON_PREFIX, typeof(RibbonButton)) + 1;
            txtLabel.Text = RibbonButton.NEW_BUTTON_PREFIX + " " + prog;
            txtName.Text = txtLabel.Text.Replace(" ", "_");
        }
        public ConfiguratoreTasto(RibbonButton btn)
        {
            InitializeComponent();

            _btn = btn;

            imgButton.Name = _btn.ImageKey;
            imgButton.Image = Utility.GetResurceImage(_btn.ImageKey);
            txtLabel.Text = _btn.Text;
            txtName.Text = _btn.Name;
            txtDesc.Text = _btn.Description;
            txtScreenTip.Text = _btn.ScreenTip;
            chkToggleButton.Checked = _btn.ToggleButton;            
            if (_btn.Slot == 1) 
            {
                radioDimSmall.Checked = true;
                ControlContainer ctrl = _btn.Parent as ControlContainer;
                if (ctrl.CtrlCount > 1)
                    radioDimLarge.Enabled = false;
            }
            else
                radioDimLarge.Checked = true;
        }

        private void ChangeBtnImage(object sender, EventArgs e)
        {
            using (CambiaImmagine chooseImageDialog = new CambiaImmagine())
            {
                if (chooseImageDialog.ShowDialog() == DialogResult.OK)
                {
                    imgButton.Name = chooseImageDialog.ResourceName;
                    imgButton.Image = chooseImageDialog.Img;
                }
            }
        }

        private void Applica_Click(object sender, EventArgs e)
        {
            if (imgButton.Name == "")
            {
                MessageBox.Show("Selezionare un'immagine prima di creare il tasto.", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            _btn.ImageKey = imgButton.Name;
            _btn.Text = txtLabel.Text;
            _btn.Name = txtName.Text;
            _btn.Description = txtDesc.Text;
            _btn.ScreenTip = txtScreenTip.Text;
            _btn.ToggleButton = chkToggleButton.Checked;
            _btn.Dimension = radioDimSmall.Checked ? 0 : 1;

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
