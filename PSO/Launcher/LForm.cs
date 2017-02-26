using System.Deployment.Application;
using System.Drawing;
using System.Windows.Forms;

namespace Iren.PSO.Launcher
{
    public partial class LForm : Form
    {
        #region Costruttore
        
        public LForm(ContextMenuStrip menu)
        {
            InitializeComponent();
#if !DEBUG
            Text = "PSO - v." + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
#endif
            int j = 0;
            foreach (ToolStripMenuItem item in menu.Items)
            {
                Button btn = new Button();
                btn.ImageList = menu.ImageList;
                btn.ImageKey = item.ImageKey;
                btn.Text = item.Text;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
                btn.Name = item.Name;
                btn.Tag = item.Tag;
                btn.Size = new Size(200, 42);
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.Margin = new Padding(0, 0, 0, 0);
                btn.Padding = new Padding(0, 0, 0, 0);
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextAlign = ContentAlignment.MiddleLeft;
                btn.Dock = DockStyle.Top;
                btn.Click += LDaemon.StartApplication;

                menuLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, btn.Height));
                menuLayout.Controls.Add(btn, 0, j++);
            }
        }
        
        #endregion

        #region Eventi
        
        private void Launcher_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }
        
        #endregion
    }
}
