using Iren.PSO.Base;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ControlliEsistenti : Form
    {
        DataTable _dtCtrl;
        Control _group;

        public ControlliEsistenti(Control group)
        {
            InitializeComponent();            
            treeViewControlli.ImageList = Utility.ImageListSmall;
            _group = group;
        }

        public ControlliEsistenti(Control group, params int[] controlType)
            : this(group)
        {
            _dtCtrl = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO);

            var typesDesc = _dtCtrl.AsEnumerable()
                .Where(r => controlType.Contains((int)r["IdTipologiaControllo"]))
                .Select(r => new { Id = r["IdTipologiaControllo"], Desc = r["DesTipologiaControllo"] })
                .Distinct();

            foreach (var type in typesDesc)
            {
                TreeNode typeRoot = new TreeNode(type.Desc.ToString());
                typeRoot.Name = type.Id.ToString();
                typeRoot.ImageIndex = 1000;
                typeRoot.SelectedImageIndex = 1000;

                var controls = _dtCtrl.AsEnumerable()
                    .Where(r => r["IdTipologiaControllo"].Equals(type.Id))
                    .Where(r => !ConfiguratoreRibbon.ControlliUtilizzati.Contains((int)r["IdControllo"]));

                foreach (var ctrl in controls)
                {
                    TreeNode c = new TreeNode(ctrl["Label"].ToString());
                    c.Tag = ctrl["IdControllo"];

                    if (!ctrl["Immagine"].Equals(""))
                    {
                        c.ImageKey = ctrl["Immagine"].ToString();
                        c.SelectedImageKey = ctrl["Immagine"].ToString();
                    }
                    else
                    {
                        c.ImageIndex = 1000;
                        c.SelectedImageIndex = 1000;
                    }
                    typeRoot.Nodes.Add(c);
                }

                if (typeRoot.Nodes.Count > 0)
                    treeViewControlli.Nodes.Add(typeRoot);
            }
            treeViewControlli.ExpandAll();

            DataTable ribbons = DataBase.Select(DataBase.SP.RIBBON.GRUPPO_CONTROLLO, "@IdApplicazione=-1;@IdUtente=-1");
            DataView functions = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_FUNZIONE).DefaultView;

            functions.RowFilter = "IdFunzione=-1";
            listBoxFunzioni.DisplayMember = "NomeFunzione";
            listBoxFunzioni.ValueMember = "IdFunzione";

            DataView gruppi = new DataView(ribbons.DefaultView.ToTable(true, "LabelGruppo", "IdGruppo", "IdControllo"));
            DataView applicazioni = new DataView(ribbons.DefaultView.ToTable(true, "IdGruppo", "DesApplicazione", "IdGruppoControllo"));

            gruppi.RowFilter = "IdControllo = -1";
            listBoxGruppi.DisplayMember = "LabelGruppo";
            listBoxGruppi.ValueMember = "IdGruppo";

            applicazioni.RowFilter = "IdGruppo = -1";
            listBoxApplicazioni.DisplayMember = "DesApplicazione";
            listBoxApplicazioni.ValueMember = "IdGruppoControllo";

            listBoxFunzioni.DataSource = functions;
            listBoxGruppi.DataSource = gruppi;
            listBoxApplicazioni.DataSource = applicazioni;

        }

        private void AfterSelectNode(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag != null && e.Node.Tag.GetType() == typeof(int))
            {
                int id = (int)e.Node.Tag;

                var selectedCtrl = _dtCtrl.AsEnumerable()
                    .Where(c => c["IdControllo"].Equals(id))
                    .FirstOrDefault();

                if (!selectedCtrl["Immagine"].Equals(""))
                {
                    imgButton.Image = Utility.ImageListNormal.Images[selectedCtrl["Immagine"].ToString()];
                    imgButton.Tag = selectedCtrl["Immagine"];
                }
                txtLabel.Text = selectedCtrl["Label"].ToString();
                txtDesc.Text = selectedCtrl["Descrizione"].ToString();
                txtScreenTip.Text = selectedCtrl["ScreenTip"].ToString();
                
                if (selectedCtrl["ControlSize"].Equals(1))
                    radioDimLarge.Checked = true;
                else if (selectedCtrl["ControlSize"].Equals(0))
                    radioDimSmall.Checked = true;
                else
                {
                    radioDimLarge.Checked = false;
                    radioDimSmall.Checked = false;
                }

                (listBoxGruppi.DataSource as DataView).RowFilter = "IdControllo = " + id;

                if(listBoxGruppi.Items.Count > 0)
                    SelectedGroupChanged(listBoxGruppi, new EventArgs());
            }
            else
            {
                radioDimLarge.Checked = false;
                radioDimSmall.Checked = false;
                imgButton.Image = null;
            }
        }

        private void SelectedGroupChanged(object sender, EventArgs e)
        {
            DataView dvG = listBoxGruppi.DataSource as DataView;
            DataView dvA = listBoxApplicazioni.DataSource as DataView;
            DataView dvF = listBoxFunzioni.DataSource as DataView;

            if (listBoxGruppi.SelectedValue != null)
            {
                dvF.RowFilter = "IdFunzione=-1";
                dvA.RowFilter = /*dvG.RowFilter + " AND*/"IdGruppo = " + listBoxGruppi.SelectedValue;

                if (listBoxApplicazioni.Items.Count > 0)
                    SelectedApplicationChanged(listBoxApplicazioni, null);
            }
            else
            {
                dvF.RowFilter = "IdFunzione=-1";
                dvA.RowFilter = "IdGruppo=-1";
            }
        }

        private void AggiungiControllo_Click(object sender, EventArgs e)
        {
            if (treeViewControlli.SelectedNode != null) 
            {
                TreeNode ctrl = treeViewControlli.SelectedNode;
                if (ctrl.Tag != null && ctrl.Tag.GetType() == typeof(int))
                {
                    ControlContainer container = Utility.CreateEmptyContainer(_group) as ControlContainer;
                    _group.Controls.Add(container);

                    if (ctrl.Parent.Name == "1" || ctrl.Parent.Name == "2")
                    {
                        RibbonButton btn = new RibbonButton(imgButton.Tag.ToString(), (int)ctrl.Tag);
                        btn.Text = txtLabel.Text;
                        btn.Dimension = radioDimLarge.Checked ? 1 : 0;
                        btn.ScreenTip = txtScreenTip.Text;
                        btn.Description = txtDesc.Text;
                        btn.ToggleButton = ctrl.Parent.Name == "2";
                        btn.Functions = listBoxFunzioni.Items
                            .Cast<DataRowView>()
                            .Select(r => (int)r["IdFunzione"])
                            .ToList();

                        container.Controls.Add(btn);
                    }
                    else
                    {
                        RibbonDropDown drp = new RibbonDropDown((int)ctrl.Tag);
                        drp.Text = txtLabel.Text;
                        drp.ScreenTip = txtScreenTip.Text;
                        drp.Description = txtDesc.Text;
                        drp.Functions = listBoxFunzioni.Items
                            .Cast<DataRowView>()
                            .Select(r => (int)r["IdFunzione"])
                            .ToList();

                        container.Controls.Add(drp);
                    }
                }
            }
        }

        private void SelectedApplicationChanged(object sender, EventArgs e)
        {
            DataView dvF = listBoxFunzioni.DataSource as DataView;
            if (listBoxApplicazioni.SelectedValue != null)
                dvF.RowFilter = "IdGruppoControllo = " + listBoxApplicazioni.SelectedValue;
            else
                dvF.RowFilter = "IdGruppoControllo = -1";
        }
    }
}
