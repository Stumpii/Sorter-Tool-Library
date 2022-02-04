using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;

namespace SorterToolLibrary.SorterTool
{
    public class HOMESheet
    {
        #region Properties

        public int Alm_SD_Max
        {
            get
            {
                if (Worksheet == null)
                    return 0;

                try
                {
                    return ExcelHelpers.ExcelCellToInt32(Worksheet.Range["ALM_SD_MAX"]);
                }
                catch
                {
                    Trace.TraceError("Unable to read named range 'ALM_SD_MAX' on HOME sheet.");
                    return 0;
                }
            }
        }

        /// <summary>
        /// The size of NVSD alarm group.
        /// </summary>
        public int NVSDCount
        {
            get
            {
                return Math.Max(VSD_Starting_Point - SD_Starting_Point, 0);
            }
        }

        public string Platform { get; private set; }

        /// <summary>
        /// Alarm number of the first NVSD
        /// </summary>
        public int SD_Starting_Point
        {
            get
            {
                if (Worksheet == null)
                    return 0;

                try
                {
                    return ExcelHelpers.ExcelCellToInt32(Worksheet.Range["SD_Starting_Point"]);
                }
                catch
                {
                    Trace.TraceError("Unable to read named range 'SD_Starting_Point' on HOME sheet.");
                    return 0;
                }
            }
        }

        public string VersionMajor { get; private set; }

        public string VersionMinor { get; private set; }

        /// <summary>
        /// The size of VSD alarm group. May be 0 is VSDs are not configured (start number = 0)
        /// </summary>
        public int VSDCount
        {
            get
            {
                return Math.Max(Alm_SD_Max - VSD_Starting_Point + 1, 0);
            }
        }

        /// <summary>
        /// Alarm number of the first VSD
        /// </summary>
        public int VSD_Starting_Point
        {
            get
            {
                if (Worksheet == null)
                    return 0;

                try
                {
                    return ExcelHelpers.ExcelCellToInt32(Worksheet.Range["VSD_Starting_Point"]);
                }
                catch
                {
                    Trace.TraceError("Unable to read named range 'VSD_Starting_Point' on HOME sheet.");
                    return 0;
                }
            }
        }

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "HOME";

        #endregion Properties

        #region Methods

        public bool ReadSheet()
        {
            try
            {
                if (Workbook != null)
                {
                    Worksheet = ((Worksheet)Workbook.Sheets[WorksheetName]);
                    return ReadSheet(Worksheet);
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                throw new ApplicationException(errMsg);
            }
        }

        public bool ReadSheet(Workbook Workbook)
        {
            try
            {
                Worksheet = ((Worksheet)Workbook.Sheets[WorksheetName]);
                return ReadSheet(Worksheet);
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                throw new ApplicationException(errMsg);
            }
        }

        public bool ReadSheet(Worksheet Worksheet)
        {
            Trace.TraceInformation($"Reading Worksheet: {Worksheet.Name}");

            try
            {
                if (Worksheet == null)
                {
                    string errMsg = $"Worksheet null";
                    return false;
                }

                bool sheetOkay = true;

                VersionMajor = ExcelHelpers.ExcelCellToString(Worksheet.Range["Ver_Major"]).Trim();
                if (string.IsNullOrWhiteSpace(VersionMajor))
                {
                    Trace.TraceError("Could not determine Major Version from the HOME sheet.");
                    sheetOkay = false;
                }

                VersionMinor = ExcelHelpers.ExcelCellToString(Worksheet.Range["Ver_Minor"]).Trim();
                if (string.IsNullOrWhiteSpace(VersionMinor))
                {
                    Trace.TraceError("Could not determine Minor Version from the HOME sheet.");
                    sheetOkay = false;
                }

                Platform = ExcelHelpers.ExcelCellToString(Worksheet.Range["Platform"]).Trim();
                if (string.IsNullOrWhiteSpace(Platform))
                {
                    Trace.TraceError("Could not determine Platform from the HOME sheet.");
                    sheetOkay = false;
                }

                return sheetOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }
}