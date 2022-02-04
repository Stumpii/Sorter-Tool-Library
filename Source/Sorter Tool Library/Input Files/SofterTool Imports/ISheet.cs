using Microsoft.Office.Interop.Excel;

namespace SorterToolLibrary.SorterTool
{
    public interface ISheet
    {
        #region Properties

        Worksheet Worksheet { get; }

        string WorksheetName { get; }

        #endregion Properties

        #region Methods

        bool ReadSheet(Workbook Workbook);

        string ToString();

        #endregion Methods
    }
}