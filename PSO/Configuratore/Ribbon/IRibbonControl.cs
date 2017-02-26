using System.Collections.Generic;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    interface IRibbonControl
    {
        int Slot { get; }
        string Description { get; }
        string Name { get; set; }
        string ImageKey { get; }
        string ScreenTip { get; set; }
        string Text { get; set; }
        bool ToggleButton { get; }
        int Dimension { get; }
        int IdTipologia { get; }
        int IdControllo { get; }
        List<int> Functions { get; set; }
        bool Enabled { get; set; }
    }
}
