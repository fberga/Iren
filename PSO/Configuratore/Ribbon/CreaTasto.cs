using System;
using System.Drawing;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class CreaTasto : Form
    {

        public string ResourceName { get; private set; }
        public Image Img { get; private set; }
        public string BtnName { get { return txtName.Text; } }
        public string BtnText { get { return txtLabel.Text; } }

        public CreaTasto(Control ribbon)
        {
            InitializeComponent();

            int prog = Utility.FindLastOfItsKind(ribbon, RibbonButton.NEW_BUTTON_PREFIX, typeof(RibbonButton)) + 1;
            txtLabel.Text = RibbonButton.NEW_BUTTON_PREFIX + " " + prog;
            txtName.Text = txtLabel.Text.Replace(" ", "_");

            imageListView.LargeImageList = Utility.ImageListNormal;

            int i = 0;
            foreach (string img in Utility.ImageListNormal.Images.Keys)
            {
                ListViewItem item = new ListViewItem();
                item.Text = img;
                item.ToolTipText = img;
                item.ImageIndex = i++;
                item.ImageKey = img;
                imageListView.Items.Add(item);
            }
        }

        private void SelectItemByDoubleClick(object sender, MouseEventArgs e)
        {
            if(e.Button == System.Windows.Forms.MouseButtons.Left) 
            {
                Applica_Click(null, null);
            }
        }

        private void Applica_Click(object sender, EventArgs e)
        {
            if (txtLabel.Text == "" || txtName.Text == "")
            {
                MessageBox.Show("Inserire un nome e/o un label per il tasto.", "ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (imageListView.SelectedItems.Count == 0)
            { 
                MessageBox.Show("Selezionare un'immagine per il tasto.", "ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            ResourceName = imageListView.SelectedItems[0].ImageKey;            
            Img = Utility.GetResurceImage(ResourceName);
            DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }

        public new DialogResult ShowDialog()
        {
            if(base.ShowDialog() == DialogResult.Cancel)
                return DialogResult.Cancel;

            if (imageListView.SelectedIndices.Count == 1)
                return DialogResult.OK;

            return DialogResult.Cancel;
        }

        private void Annulla_Click(object sender, EventArgs e)
        {
            imageListView.SelectedIndices.Clear();
            
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            
            Close();
        }
    }
}
