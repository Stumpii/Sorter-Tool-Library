using Microsoft.Office.Interop.Excel;

namespace D_R_Engine.SorterTool
{
    public abstract class SorterToolSheet
    {
        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; set; }

        public abstract string WorksheetName { get; }

        public override string ToString()
        {
            return WorksheetName;
        }
    }

  


}