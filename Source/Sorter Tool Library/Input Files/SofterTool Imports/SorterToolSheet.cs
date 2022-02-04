using Microsoft.Office.Interop.Excel;

namespace SorterToolLibrary.SorterTool
{
    public abstract class SorterToolSheet
    {
        #region Properties

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; set; }

        public abstract string WorksheetName { get; }

        #endregion Properties

        #region Methods

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }
}