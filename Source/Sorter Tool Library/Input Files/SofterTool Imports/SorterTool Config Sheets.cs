using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Diagnostics;

namespace SorterToolLibrary.SorterTool
{
    public class ALM_GENSheet : ISheet
    {
        #region Fields

        public class ALM_GENRow
        {
            private ALM_GENSheet Parent;

            public ALM_GENRow(ALM_GENSheet aLM_GENSheet)
            {
                this.Parent = aLM_GENSheet;
            }

            #region Properties

            public string AlarmClass { get; set; }

            /// <summary>
            /// Alarm number of the shutdown, typically starting at 513
            /// </summary>
            public int? AlarmNumber { get; set; }

            public string AlarmType { get; set; }

            public int? AnlgStatIndex { get; set; }

            public int? BentlyStatus { get; set; }

            public int? Index { get; set; }

            public string IOType { get; set; }

            public string VotingGroup { get; set; }

            public ALM_INSheet.ALM_INRow ALM_IN
            {
                get
                {
                    // Skip null and 0 alarm numbers
                    if ((AlarmNumber ?? 0) == 0)
                        return null;

                    ALM_INSheet.ALM_INRow alarmIn = Parent.Parent.FindALM_INRow(AlarmNumber);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in ALM_IN for {this}.");
                        return null;
                    }
                }
            }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row[(int)ExcelCol.A]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.A]);

                if (Index == default(int?) || Index == 0) return false;

                AlarmNumber = Convert.IsDBNull(row[(int)ExcelCol.A]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.A]);
                AnlgStatIndex = Convert.IsDBNull(row[(int)ExcelCol.C]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.C]);
                IOType = Convert.IsDBNull(row[(int)ExcelCol.D]) ? "" : Convert.ToString(row[(int)ExcelCol.D]);
                BentlyStatus = Convert.IsDBNull(row[(int)ExcelCol.E]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.E]);
                AlarmType = Convert.IsDBNull(row[(int)ExcelCol.G]) ? "" : Convert.ToString(row[(int)ExcelCol.G]);
                AlarmClass = Convert.IsDBNull(row[(int)ExcelCol.M]) ? "" : Convert.ToString(row[(int)ExcelCol.M]);
                VotingGroup = Convert.IsDBNull(row[(int)ExcelCol.T]) ? "" : Convert.ToString(row[(int)ExcelCol.T]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try

                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    AlarmNumber = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);
                    AnlgStatIndex = ExcelHelpers.ExcelCellToInt32(Row.Range["C1"]);
                    IOType = ExcelHelpers.ExcelCellToString(Row.Range["D1"]);
                    BentlyStatus = ExcelHelpers.ExcelCellToInt32(Row.Range["E1"]);
                    AlarmType = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    AlarmClass = Row.Range["M1"].Value?.ToString().Trim() ?? "";
                    VotingGroup = Row.Range["T1"].Value?.ToString().Trim() ?? "";

                    return true;
                }
                catch (Exception ex)
                {
                    string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                    Trace.TraceError(errMsg);
                    throw new ApplicationException(errMsg);
                }
            }

            public override string ToString()
            {
                return $"ALM{AlarmNumber:000}";
            }

            #endregion Methods
        }

        private readonly long firstDataRow = 4;
        private SorterToolImporter Parent;

        public ALM_GENSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        #endregion Fields

        #region Properties

        internal List<ALM_GENRow> Rows { get; set; } = new List<ALM_GENRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "ALM_GEN";

        #endregion Properties

        #region Methods

        public bool ReadDataTable(DataSet DataSet)
        {
            return ReadDataTable(DataSet.Tables[WorksheetName]);
        }

        public bool ReadDataTable(DataTable DataTable)
        {
            // Read sheet
            Rows.Clear();
            foreach (DataRow row in DataTable.Rows)
            {
                var item = new ALM_GENRow(this);
                if (item.ImportRow(row))
                    Rows.Add(item);
            }

            return true;
        }

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

                Rows = new List<ALM_GENRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        ALM_GENRow item = new ALM_GENRow(this);
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return true;
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

    public class SD_GENSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 4;
        private SorterToolImporter Parent;

        public SD_GENSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        public class SD_GENRow
        {
            private SD_GENSheet Parent;

            public SD_GENRow(SD_GENSheet sD_GENSheet)
            {
                this.Parent = sD_GENSheet;
            }

            #region Properties

            public string AlarmClass { get; set; }

            /// <summary>
            /// Alarm number of the shutdown, typically starting at 513
            /// </summary>
            public int? AlarmNumber { get; set; }

            public string AlarmType { get; set; }

            public int? AnlgStatIndex { get; set; }

            public int? BentlyStatus { get; set; }

            /// <summary>
            /// Index into the shutdowns, starting at 1
            /// </summary>
            public int? Index { get; set; }

            public string IOType { get; set; }

            public bool IsVented
            {
                get
                {
                    return (VentedOrNonVented == "SHUTDOWNS_V")
                      || (VentedOrNonVented == "VSD");
                }
            }

            /// <summary>
            /// Either "VSD" or "NVSD".
            /// </summary>
            public string VentedOrNonVented { get; set; }

            public string VotingGroup { get; set; }

            public SD_INSheet.SD_INRow SD_IN
            {
                get
                {
                    // Skip null and 0 alarm numbers
                    if ((AlarmNumber ?? 0) == 0)
                        return null;

                    SD_INSheet.SD_INRow alarmIn = Parent.Parent.FindSD_INRow(AlarmNumber);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in SD_IN for {this}.");
                        return null;
                    }
                }
            }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row[(int)ExcelCol.A]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.A]);

                if (Index == default(int?) || Index == 0) return false;

                AnlgStatIndex = Convert.IsDBNull(row[(int)ExcelCol.C]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.C]);
                IOType = Convert.IsDBNull(row[(int)ExcelCol.D]) ? "" : Convert.ToString(row[(int)ExcelCol.D]);
                BentlyStatus = Convert.IsDBNull(row[(int)ExcelCol.E]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.E]);
                AlarmType = Convert.IsDBNull(row[(int)ExcelCol.G]) ? "" : Convert.ToString(row[(int)ExcelCol.G]);
                AlarmClass = Convert.IsDBNull(row[(int)ExcelCol.M]) ? "" : Convert.ToString(row[(int)ExcelCol.M]);
                VentedOrNonVented = Convert.IsDBNull(row[(int)ExcelCol.O]) ? "" : Convert.ToString(row[(int)ExcelCol.O]);
                AlarmNumber = Convert.IsDBNull(row[(int)ExcelCol.AC]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.AC]);
                VotingGroup = Convert.IsDBNull(row[(int)ExcelCol.AD]) ? "" : Convert.ToString(row[(int)ExcelCol.AD]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    AnlgStatIndex = ExcelHelpers.ExcelCellToInt32(Row.Range["C1"]);
                    IOType = ExcelHelpers.ExcelCellToString(Row.Range["D1"]);
                    BentlyStatus = ExcelHelpers.ExcelCellToInt32(Row.Range["E1"]);
                    AlarmType = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    AlarmClass = Row.Range["M1"].Value?.ToString().Trim() ?? "";
                    VentedOrNonVented = Row.Range["O1"].Value?.ToString().Trim() ?? "";
                    AlarmNumber = ExcelHelpers.ExcelCellToInt32(Row.Range["AC1"]);
                    VotingGroup = Row.Range["AD1"].Value?.ToString().Trim() ?? "";

                    return true;
                }
                catch (Exception ex)
                {
                    string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                    Trace.TraceError(errMsg);
                    throw new ApplicationException(errMsg);
                }
            }

            public override string ToString()
            {
                return $"ALM{AlarmNumber:000}";
            }

            #endregion Methods
        }

        #endregion Fields

        #region Properties

        internal List<SD_GENRow> Rows { get; set; } = new List<SD_GENRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "SD_GEN";

        #endregion Properties

        #region Methods

        public bool ReadDataTable(DataSet DataSet)
        {
            return ReadDataTable(DataSet.Tables[WorksheetName]);
        }

        public bool ReadDataTable(DataTable DataTable)
        {
            // Read sheet
            Rows.Clear();
            foreach (DataRow row in DataTable.Rows)
            {
                var item = new SD_GENRow(this);
                if (item.ImportRow(row))
                    Rows.Add(item);
            }

            return true;
        }

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

                Rows = new List<SD_GENRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        SD_GENRow item = new SD_GENRow(this);
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return true;
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

    public class STD_DEV_SEL_3Sheet : ISheet
    {
        #region Fields

        public class STD_DEV_SEL_3Row
        {
            #region Properties

            public int? Index { get; set; }

            public int? SEL_ANLG { get; set; }

            public string SEL_MODE { get; set; }

            public string SF_DIR { get; set; }

            public string VotingGroup { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["Number"]) ? default(int) : Convert.ToInt32(row["Number"]);

                if (Index == default(double) || Index == 0) return false;

                VotingGroup = Convert.IsDBNull(row["VotingGroup"]) ? default(string) : Convert.ToString(row["VotingGroup"]);
                SEL_ANLG = Convert.IsDBNull(row["SEL_ANLG"]) ? default(int) : Convert.ToInt32(row["SEL_ANLG"]);
                SF_DIR = Convert.IsDBNull(row["SF_DIR"]) ? default(string) : Convert.ToString(row["SF_DIR"]);
                SEL_MODE = Convert.IsDBNull(row["SEL_MODE"]) ? default(string) : Convert.ToString(row["SEL_MODE"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    // TODO - The header of this sheet is all messed up. Needs to be fixed to read the data correctly.
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    VotingGroup = ExcelHelpers.ExcelCellToString(Row.Range["B1"]).Trim();
                    SEL_ANLG = ExcelHelpers.ExcelCellToInt32(Row.Range["C1"]);
                    SEL_MODE = Row.Range["D1"].Value?.ToString().Trim() ?? "";
                    SF_DIR = Row.Range["E1"].Value?.ToString().Trim() ?? "";

                    return true;
                }
                catch (Exception ex)
                {
                    string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                    Trace.TraceError(errMsg);
                    throw new ApplicationException(errMsg);
                }
            }

            public override string ToString()
            {
                return $"Index: {Index:000}";
            }

            #endregion Methods
        }

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        internal List<STD_DEV_SEL_3Row> Rows { get; set; } = new List<STD_DEV_SEL_3Row>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "STD_DEV_SEL_3";

        #endregion Properties

        #region Methods

        public bool ReadDataTable(DataSet DataSet)
        {
            return ReadDataTable(DataSet.Tables[WorksheetName]);
        }

        public bool ReadDataTable(DataTable DataTable)
        {
            // Read sheet
            Rows.Clear();
            foreach (DataRow row in DataTable.Rows)
            {
                var item = new STD_DEV_SEL_3Row();
                if (item.ImportRow(row))
                    Rows.Add(item);
            }

            return true;
        }

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

                Rows = new List<STD_DEV_SEL_3Row>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        STD_DEV_SEL_3Row item = new STD_DEV_SEL_3Row();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return true;
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