using System;
using System.Drawing;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class CambiaImmagine : Form
    {
        public string ResourceName { get; private set; }
        public Image Img { get; private set; }

        public CambiaImmagine()
        {
            InitializeComponent();

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

        private void Applica_Click(object sender, EventArgs e)
        {
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

        private void SelectItemByDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                Applica_Click(null, null);
            }
        }
    }
}
