using System.Collections.Generic;

namespace IEOperateCore.Common
{
    public class HtmlTable
    {
        public List<HtmlTableRow> HtmlTableRows;

        public HtmlTable()
        {
            HtmlTableRows = new List<HtmlTableRow>();
        }
    }

    public class HtmlTableRow
    {
        public HtmlTableRow()
        {
            Cells = new List<HtmlTableCell>();
            RowIndex = -1;
        }
        public List<HtmlTableCell> Cells { get; set; }

        public int RowIndex { get; set; }
    }

    public class HtmlTableCell
    {
        public HtmlTableCell()
        {
            CellColumnIndex = -1;
            CellName = null;
            CellValue = null;
        }
        public int CellColumnIndex { get; set; }

        public string CellName { get; set; }

        public string CellValue { get; set; }
    }

}
