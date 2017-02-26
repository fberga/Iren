using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class AssegnaFunzioni : Form
    {
        IRibbonControl _ctrl;

        public AssegnaFunzioni(IRibbonControl ctrl, RibbonGroup grp, int appID, int usrID)
        {
            InitializeComponent();
            _ctrl = ctrl;

            DataTable allFunctions = DataBase.Select(DataBase.SP.RIBBON.FUNZIONE);
            foreach (int idFunzione in ctrl.Functions)
            {
                var func =
                    (from r in allFunctions.AsEnumerable()
                        where r["IdFunzione"].Equals(idFunzione)
                        select r).First();

                TreeNode f = new TreeNode();
                f.Text = func["NomeFunzione"].ToString();
                f.Tag = func["IdFunzione"];

                if (!treeViewUtilized.Nodes.ContainsKey(func["Evento"].ToString()))
                {
                    TreeNode evento = new TreeNode(func["Evento"].ToString());
                    evento.Name = func["Evento"].ToString();
                    treeViewUtilized.Nodes.Add(evento);
                }

                treeViewUtilized.Nodes[func["Evento"].ToString()].Nodes.Add(f);
                allFunctions.Rows.Remove(func);
            }

            foreach (DataRow func in allFunctions.Rows)
            {
                TreeNode f = new TreeNode();
                f.Text = func["NomeFunzione"].ToString();
                f.Tag = func["IdFunzione"];

                if (!treeViewNotUtilized.Nodes.ContainsKey(func["Evento"].ToString()))
                {
                    TreeNode evento = new TreeNode(func["Evento"].ToString());
                    evento.Name = func["Evento"].ToString();
                    treeViewNotUtilized.Nodes.Add(evento);
                }

                treeViewNotUtilized.Nodes[func["Evento"].ToString()].Nodes.Add(f);
            }
        }
        private void AggiungiFunzione_Click(object sender, EventArgs e)
        {
            TreeNode selected = treeViewNotUtilized.SelectedNode;
            if (selected != null && selected.Tag != null)
            {
                string pName = selected.Parent.Name;
                if (!treeViewUtilized.Nodes.ContainsKey(pName))
                {
                    TreeNode evento = new TreeNode(pName);
                    evento.Name = pName;
                    treeViewUtilized.Nodes.Add(evento);
                }
                if (treeViewNotUtilized.Nodes[pName].Nodes.Count == 1)
                    treeViewNotUtilized.Nodes.Remove(selected.Parent);
                else
                    treeViewNotUtilized.Nodes[pName].Nodes.Remove(selected);

                treeViewUtilized.Nodes[pName].Nodes.Add(selected);
            }
        }
        private void RimuoviFunzione_Click(object sender, EventArgs e)
        {
            TreeNode selected = treeViewUtilized.SelectedNode;
            if (selected != null && selected.Tag != null)
            {
                string pName = selected.Parent.Name;
                if (!treeViewNotUtilized.Nodes.ContainsKey(pName))
                {
                    TreeNode evento = new TreeNode(pName);
                    evento.Name = pName;
                    treeViewNotUtilized.Nodes.Add(evento);
                }
                if (treeViewUtilized.Nodes[pName].Nodes.Count == 1)
                    treeViewUtilized.Nodes.Remove(selected.Parent);
                else
                    treeViewUtilized.Nodes[pName].Nodes.Remove(selected);
                
                treeViewNotUtilized.Nodes[pName].Nodes.Add(selected);
            }
        }
        private void AssegnaFunzioni_Click(object sender, EventArgs e)
        {
            _ctrl.Functions.Clear();
            foreach (TreeNode FunctionType in treeViewUtilized.Nodes)
            {
                _ctrl.Functions = new List<int>();
                foreach (TreeNode function in FunctionType.Nodes)
                    _ctrl.Functions.Add((int)function.Tag);
            }

            DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }
        private void AnnullaCambiamenti_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            Close();
        }

        private void SpostaSotto_Click(object sender, EventArgs e)
        {
            if (treeViewUtilized.SelectedNode != null && treeViewUtilized.SelectedNode.Tag != null
                && treeViewUtilized.SelectedNode.NextNode != null && treeViewUtilized.SelectedNode.NextNode.Tag != null)
            {
                int i = treeViewUtilized.SelectedNode.Index;
                TreeNode n = treeViewUtilized.SelectedNode;
                TreeNode parent = treeViewUtilized.SelectedNode.Parent;
                
                treeViewUtilized.SelectedNode.Remove();
                parent.Nodes.Insert(i + 1, n);
                treeViewUtilized.SelectedNode = n;
            }
        }
        private void SpostaSopra_Click(object sender, EventArgs e)
        {
            if (treeViewUtilized.SelectedNode != null && treeViewUtilized.SelectedNode.Tag != null
                && treeViewUtilized.SelectedNode.PrevNode != null && treeViewUtilized.SelectedNode.PrevNode.Tag != null)
            {
                int i = treeViewUtilized.SelectedNode.Index;
                TreeNode n = treeViewUtilized.SelectedNode;
                TreeNode parent = treeViewUtilized.SelectedNode.Parent;

                treeViewUtilized.SelectedNode.Remove();
                parent.Nodes.Insert(i - 1, n);
                treeViewUtilized.SelectedNode = n;
            }
        }
    }
}
