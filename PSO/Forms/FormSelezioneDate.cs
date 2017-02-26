using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Iren.PSO.Forms
{
    public partial class FormSelezioneDate : Form
    {
        #region Variabili

        private List<DateTime> _outList = new List<DateTime>();
        private Dictionary<DateTime, bool> _workList = new Dictionary<DateTime, bool>();
        private Dictionary<DateTime, bool> _workListOld = new Dictionary<DateTime, bool>();
        private Dictionary<Tuple<DateTime, DateTime>, bool> _clusters = new Dictionary<Tuple<DateTime, DateTime>, bool>();
        private DateTime _extraDateFrom = new DateTime();

        #endregion

        #region Proprietà

        public List<DateTime> SelectedDates { get { return _outList; } }

        #endregion

        #region Costruttori

        public FormSelezioneDate()
        {
            InitializeComponent();
            
            if (Struct.intervalloGiorni > 0)
            {
                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
                categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL AND IdApplicazione = " + Workbook.IdApplicazione;
                entitaProprieta.RowFilter = "IdApplicazione = " + Workbook.IdApplicazione;

                _extraDateFrom = Workbook.DataAttiva.AddDays(Struct.intervalloGiorni + 1);

                SortedList<DateTime, string> giorniExtra = new SortedList<DateTime, string>();
                int maxIntervallo = Struct.intervalloGiorni;
                foreach (DataRowView entita in categoriaEntita)
                {
                    entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta LIKE '%GIORNI_STRUTTURA' AND IdApplicazione = " + Workbook.IdApplicazione;
                    if (entitaProprieta.Count > 0)
                    {
                        int value = int.Parse(entitaProprieta[0]["Valore"].ToString());
                        maxIntervallo = Math.Max(maxIntervallo, value);
                        if (value > Struct.intervalloGiorni)
                            if (giorniExtra.ContainsKey(Workbook.DataAttiva.AddDays(value)))
                                giorniExtra[Workbook.DataAttiva.AddDays(value)] += ", " + entita["SiglaEntita"].ToString().Replace("UP_", "");
                            else
                                giorniExtra.Add(Workbook.DataAttiva.AddDays(value), entita["SiglaEntita"].ToString().Replace("UP_", ""));
                    }
                }

                if (Struct.intervalloGiorni > 0)
                {
                    _clusters.Add(Tuple.Create(Workbook.DataAttiva, giorniExtra.Count > 0 ? giorniExtra.Last().Key : Workbook.DataAttiva.AddDays(Struct.intervalloGiorni)), false);
                    checkClusterDate.Items.Add("Tutti");
                }
                if (giorniExtra.Count > 0)
                {
                    _clusters.Add(Tuple.Create(Workbook.DataAttiva, Workbook.DataAttiva.AddDays(Struct.intervalloGiorni)), false);
                    checkClusterDate.Items.Add("Da " + Workbook.DataAttiva.ToString("ddd dd MMM") + " a " + Workbook.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("ddd dd MMM"));

                    foreach (var kv in giorniExtra)
                    {
                        _clusters.Add(Tuple.Create(Workbook.DataAttiva.AddDays(Struct.intervalloGiorni + 1), kv.Key), false);
                        checkClusterDate.Items.Add("Da " + Workbook.DataAttiva.AddDays(Struct.intervalloGiorni + 1).ToString("ddd dd MMM") + " a " + kv.Key.ToString("ddd dd MMM") + " (" + kv.Value + ")");
                    }
                }

                for (int i = 0; i <= maxIntervallo; i++)
                {
                    _workList.Add(Workbook.DataAttiva.AddDays(i), false);
                    checkDate.Items.Add(Workbook.DataAttiva.AddDays(i).ToString("dddd d MMMM yyyy"));
                }                

                panelTop.FixedPanel = FixedPanel.None;
                panelTop.SplitterDistance = checkClusterDate.GetItemRectangle(0).Height * checkClusterDate.Items.Count + 8;
                panelTop.FixedPanel = FixedPanel.Panel1;

                this.Height = panelTop.SplitterDistance + checkDate.GetItemRectangle(0).Height * (checkDate.Items.Count + 1) + 3 + panelButtons.Height;
            }
        }

        #endregion

        #region Eventi

        private void CheckedListBox_MouseMove(object sender, MouseEventArgs e)
        {
            CheckedListBox chk = (CheckedListBox)sender;

            Point point = chk.PointToClient(Cursor.Position);
            int index = chk.IndexFromPoint(point);
            if (index < 0) return;

            chk.SelectedItem = chk.Items[index];
        }
        private void CheckedListBox_MouseLeave(object sender, EventArgs e)
        {
            ((CheckedListBox)sender).SelectedItem = null;
        }
        private void checkClusterDate_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            checkClusterDate.ItemCheck -= checkClusterDate_ItemCheck;

            var range = _clusters.ElementAt(e.Index).Key;

            if (e.NewValue == CheckState.Checked)
            {
                var toCheck = _clusters.Select((kv, i) => new { Index = i, Key = kv.Key }).Where(kv => kv.Key.Item1 >= range.Item1 && kv.Key.Item2 <= range.Item2);
                foreach (var chk in toCheck.ToList())
                {
                    _clusters[chk.Key] = true;
                    checkClusterDate.SetItemChecked(chk.Index, true);
                }
            }
            else
            {
                var toUncheck = _clusters.Select((kv, i) => new { Index = i, Key = kv.Key }).Where(kv => (kv.Key.Item1 >= range.Item1 && kv.Key.Item1 <= range.Item2) || (kv.Key.Item2 >= range.Item1 && kv.Key.Item2 <= range.Item2));
                foreach (var chk in toUncheck.ToList())
                {
                    _clusters[chk.Key] = false;
                    checkClusterDate.SetItemChecked(chk.Index, false);
                }
            }
            
            //reset
            for (int i = 0; i < checkDate.Items.Count; i++)
                checkDate.SetItemChecked(i, false);

            var datesToCheck = _clusters.Where(kv => kv.Value).Select(kv => kv.Key);
            foreach (var date in datesToCheck)
            {
                var toCheck = _workList.Select((kv, i) => new { Index = i, Value = kv }).Where((val) => val.Value.Key >= date.Item1 && val.Value.Key <= date.Item2);
                foreach (var check in toCheck.ToList())
                    checkDate.SetItemChecked(check.Index, true);
            }
            

            checkClusterDate.ItemCheck += checkClusterDate_ItemCheck;
        }
        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            _workList = new Dictionary<DateTime,bool>(_workListOld);

            checkDate.ItemCheck -= checkDate_ItemCheck;

            for (int i = 0; i < _workList.Count; i++)
                checkDate.SetItemChecked(i, _workList.ElementAt(i).Value);

            checkDate.ItemCheck += checkDate_ItemCheck;

            this.Hide();
        }
        private void btnApplica_Click(object sender, EventArgs e)
        {
            _outList =
                (from kv in _workList
                 where kv.Value
                 select kv.Key).ToList();

            this.Hide();
        }
        private void checkDate_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            _workList[_workList.ElementAt(e.Index).Key] = e.NewValue == CheckState.Checked;
        }
        private void FormSelezioneDate_VisibleChanged(object sender, EventArgs e)
        {
            if (this.Visible)
                _workListOld = new Dictionary<DateTime, bool>(_workList);
        }

        public void SelectAll()
        {
            if(checkClusterDate.Items.Count > 0)
                checkClusterDate.SetItemChecked(0, true);
        }

        #endregion
    }
}
