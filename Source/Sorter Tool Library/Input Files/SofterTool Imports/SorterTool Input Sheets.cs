using SorterToolLibrary.SorterTool;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SorterToolLibrary.SorterTool
{
    public class AI_INSheet : SorterToolSheet, ISheet
    {
        #region Fields

        private static readonly long firstDataRow = 2;
        private SorterToolImporter Parent;

        #endregion Fields

        public AI_INSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        #region Properties

        public List<AI_INRow> Rows { get; set; } = new List<AI_INRow>();

        public override string WorksheetName => "AI_IN";

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
                var item = new AI_INRow(this);
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

                Rows = new List<AI_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        AI_INRow item = new AI_INRow(this);
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

        #endregion Methods

        #region Classes

        public class AI_INRow
        {
            #region Fields

            private AI_INSheet Parent;

            #endregion Fields

            #region Constructors

            public AI_INRow(AI_INSheet aI_INSheet)
            {
                this.Parent = aI_INSheet;
            }

            #endregion Constructors

            #region Properties

            public string Address { get; set; }

            /// <summary>
            /// Analog number (i.e. 1 to 256)
            /// </summary>
            public int? Anlg { get; private set; }

            public string Description { get; set; }

            public int? Index { get; set; }

            public string LocationOnNetwork { get; set; }

            /// <summary>
            /// Module point reference (i.e. AI01_CH01)
            /// </summary>
            public string ModulePoint { get; set; }

            public string SignalType { get; set; }

            public string SlotNumber { get; set; }

            public string Specifier { get; set; }

            public string ClientTag { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["#"]) ? default(int?) : Convert.ToInt32(row["#"]);

                if (Index == default || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                Anlg = Convert.IsDBNull(row["ANLG#"]) ? default(int?) : Convert.ToInt32(row["ANLG#"]);
                LocationOnNetwork = Convert.IsDBNull(row["Location on Network"]) ? "" : Convert.ToString(row["Location on Network"]);
                SlotNumber = Convert.IsDBNull(row["Slot #"]) ? "" : Convert.ToString(row["Slot #"]);
                ModulePoint = Convert.IsDBNull(row["Module Point"]) ? "" : Convert.ToString(row["Module Point"]);
                ClientTag = Convert.IsDBNull(row["Client Tag"]) ? "" : Convert.ToString(row["Client Tag"]);
                SignalType = Convert.IsDBNull(row["Signal Type"]) ? "" : Convert.ToString(row["Signal Type"]);
                Address = Convert.IsDBNull(row["Address"]) ? "" : Convert.ToString(row["Address"]);
                Specifier = Convert.IsDBNull(row["SPECIFIER"]) ? "" : Convert.ToString(row["SPECIFIER"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default || Index == 0) return false;

                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["D1"]).Trim();
                    LocationOnNetwork = Row.Range["B1"].Value?.ToString().Trim() ?? "";
                    SlotNumber = Row.Range["C1"].Value?.ToString().Trim() ?? "";
                    Anlg = ExcelHelpers.ExcelCellToInt32(Row.Range["E1"]);
                    ClientTag = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    Description = Row.Range["H1"].Value?.ToString().Trim() ?? "";
                    SignalType = Row.Range["I1"].Value?.ToString().Trim() ?? "";
                    Address = Row.Range["M1"].Value?.ToString().Trim() ?? "";
                    Specifier = Row.Range["N1"].Value?.ToString().Trim() ?? "";

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
                return ModulePoint;
            }

            #endregion Methods

            public I_O_LISTSheet.I_O_LISTRow I_O_List
            {
                get
                {
                    I_O_LISTSheet.I_O_LISTRow alarmIn = Parent.Parent.FindI_O_LISTRow(ModulePoint);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in I_O_LIST for {this}.");
                        return null;
                    }
                }
            }
        }

        #endregion Classes
    }

    public class ALM_INSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        public List<ALM_INRow> Rows { get; set; } = new List<ALM_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "ALM_IN";

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
                var item = new ALM_INRow();
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

                Rows = new List<ALM_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        ALM_INRow item = new ALM_INRow();
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

        #region Classes

        public class ALM_INRow
        {
            #region Properties

            public string AlarmDescription { get; set; }

            public int? AlarmNumber { get; private set; }

            public string AlarmRef { get; set; }

            public string AlarmTag { get; set; }

            /// <summary>
            /// Alarm type (i.e. L, H)
            /// </summary>
            public string AlarmType { get; set; }
            public string CustomPLCTag { get; set; }
            public int? Index { get; set; }

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

                if (Index == default || Index == 0) return false;

                AlarmNumber = Convert.IsDBNull(row["Alarm Number"]) ? default(int?) : Convert.ToInt32(row["Alarm Number"]);
                AlarmRef = Convert.IsDBNull(row["Alarm Ref"]) ? "" : Convert.ToString(row["Alarm Ref"]);
                AlarmDescription = Convert.IsDBNull(row["Alarm Description"]) ? "" : Convert.ToString(row["Alarm Description"]);
                AlarmTag = Convert.IsDBNull(row["Alarm Tag"]) ? "" : Convert.ToString(row["Alarm Tag"]);
                AlarmType = Convert.IsDBNull(row["Alarm Type"]) ? "" : Convert.ToString(row["Alarm Type"]);
                CustomPLCTag = Convert.IsDBNull(row["CustomPLCTag"]) ? "" : Convert.ToString(row["CustomPLCTag"]);
                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default || Index == 0) return false;

                    AlarmRef = ExcelHelpers.ExcelCellToString(Row.Range["B1"]).Trim();
                    AlarmDescription = Row.Range["C1"].Value?.ToString().Trim() ?? "";

                    // In AB Sorter Tool, G1 is used instead of F1
                    if ((Row.Range["G1"].Value?.ToString() ?? "") != "")
                        AlarmNumber = ExcelHelpers.ExcelCellToInt32(Row.Range["G1"]);
                    else
                        AlarmNumber = ExcelHelpers.ExcelCellToInt32(Row.Range["F1"]);

                    AlarmTag = Row.Range["J1"].Value?.ToString().Trim() ?? "";
                    AlarmType = Row.Range["L1"].Value?.ToString().Trim() ?? "";

                    CustomPLCTag = ExcelHelpers.ExcelCellToString(Row.Range["U1"]).Trim();
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
                return AlarmRef;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class ANLG_INSheet : SorterToolSheet, ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly int headerRow = 1;

        #endregion Fields

        #region Properties

        public List<ANLG_INRow> Rows { get; set; } = new List<ANLG_INRow>();

        public override string WorksheetName => "ANLG_IN";

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
                var item = new ANLG_INRow();
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

                Rows = new List<ANLG_INRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        ANLG_INRow item = new ANLG_INRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        ANLG_INRow item = new ANLG_INRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
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

        #region Classes

        public class ANLG_INRow
        {
            #region Properties

            public string Address { get; set; }

            public static int AddressIdx { get; private set; }

            public string ANLGNum { get; private set; }

            public static int ANLGNumIdx { get; private set; }

            public string ClientTag { get; set; }

            public static int ClientTagIdx { get; private set; }

            public int ConversionFormula { get; private set; }

            public static int ConversionFormulaIdx { get; private set; }

            public string Description { get; set; }

            public static int DescriptionIdx { get; private set; }

            public double? EngMax { get; set; }

            public double? EngMin { get; set; }

            public double? HiAlmLimit { get; set; }

            public static int HiAlmLimitIdx { get; private set; }

            public double? HiSDLimit { get; set; }

            public static int HiSDLimitIdx { get; private set; }

            public int? Index { get; set; }

            public static int ITEMNumIdx { get; private set; }

            public double? LoAlmLimit { get; set; }

            public static int LoAlmLimitIdx { get; private set; }

            public double? LoSDLimit { get; set; }

            public static int LoSDLimitIdx { get; private set; }

            public static int MaxIdx { get; private set; }

            public static int MinIdx { get; private set; }

            public double? OpenLimit { get; set; }

            public static int OpenLimitIdx { get; private set; }

            public double? SFHiLimit { get; set; }

            public static int SFHiLimitIdx { get; private set; }

            public double? SFLoLimit { get; set; }

            public static int SFLoLimitIdx { get; private set; }

            public string Source { get; private set; }

            public static int SourceIdx { get; private set; }

            public string Unit { get; private set; }

            public static int UnitIdx { get; private set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["ITEM#"]) ? default(int?) : Convert.ToInt32(row["ITEM#"]);

                if (Index == default(int?) || Index == 0) return false;

                ANLGNum = Convert.IsDBNull(row["ANLG#"]) ? "" : Convert.ToString(row["ANLG#"]);
                ClientTag = Convert.IsDBNull(row["ClientTag"]) ? "" : Convert.ToString(row["ClientTag"]);
                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                Address = Convert.IsDBNull(row["Address"]) ? "" : Convert.ToString(row["Address"]);
                EngMin = Convert.IsDBNull(row["Min."]) ? 0.0 : Convert.ToDouble(row["Min."]);
                EngMax = Convert.IsDBNull(row["Max."]) ? 0.0 : Convert.ToDouble(row["Max."]);

                SFLoLimit = Convert.IsDBNull(row["SFLoLimit"]) ? 0.0 : Convert.ToDouble(row["SFLoLimit"]);
                SFHiLimit = Convert.IsDBNull(row["SFHiLimit"]) ? 0.0 : Convert.ToDouble(row["SFHiLimit"]);
                LoAlmLimit = Convert.IsDBNull(row["LoAlmLimit"]) ? 0.0 : Convert.ToDouble(row["LoAlmLimit"]);
                LoSDLimit = Convert.IsDBNull(row["LoSDLimit"]) ? 0.0 : Convert.ToDouble(row["LoSDLimit"]);
                HiAlmLimit = Convert.IsDBNull(row["HiAlmLimit"]) ? 0.0 : Convert.ToDouble(row["HiAlmLimit"]);
                HiSDLimit = Convert.IsDBNull(row["HiSDLimit"]) ? 0.0 : Convert.ToDouble(row["HiSDLimit"]);
                OpenLimit = Convert.IsDBNull(row["OpenLimit"]) ? 0.0 : Convert.ToDouble(row["OpenLimit"]);
                Source = Convert.IsDBNull(row["Source"]) ? "" : Convert.ToString(row["Source"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Columns[ITEMNumIdx]);

                    if (Index == default(int?) || Index == 0) return false;

                    Address = ExcelHelpers.ExcelCellToString(Row.Columns[AddressIdx]);
                    ANLGNum = ExcelHelpers.ExcelCellToString(Row.Columns[ANLGNumIdx]);
                    ClientTag = ExcelHelpers.ExcelCellToString(Row.Columns[ClientTagIdx]);
                    Description = ExcelHelpers.ExcelCellToString(Row.Columns[DescriptionIdx]);
                    Unit = ExcelHelpers.ExcelCellToString(Row.Columns[UnitIdx]);
                    EngMin = ExcelHelpers.ExcelCellToDouble(Row.Columns[MinIdx]);
                    EngMax = ExcelHelpers.ExcelCellToDouble(Row.Columns[MaxIdx]);
                    SFLoLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[SFLoLimitIdx]);
                    SFHiLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[SFHiLimitIdx]);
                    LoAlmLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[LoAlmLimitIdx]);
                    LoSDLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[LoSDLimitIdx]);
                    HiAlmLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[HiAlmLimitIdx]);
                    HiSDLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[HiSDLimitIdx]);
                    OpenLimit = ExcelHelpers.ExcelCellToDouble(Row.Columns[OpenLimitIdx]);
                    ConversionFormula = ExcelHelpers.ExcelCellToInt32(Row.Columns[ConversionFormulaIdx]);
                    // TODO - Determine if Source of ANLG_IN is needed.
                    //Source = ExcelHelpers.ExcelCellToString(Row.Columns[SourceIdx]);

                    return true;
                }
                catch (Exception ex)
                {
                    string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                    Trace.TraceError(errMsg);
                    throw new ApplicationException(errMsg);
                }
            }

            public bool ReadColumnHeaders(Range Row)
            {
                int result = 0;
                ITEMNumIdx = ExcelHelpers.FindColumnHeader(Row, "ITEM#");
                result = Math.Min(result, ITEMNumIdx); // Result will be -1 if column does not exist.
                AddressIdx = ExcelHelpers.FindColumnHeader(Row, "Address");
                result = Math.Min(result, AddressIdx); // Result will be -1 if column does not exist.
                ANLGNumIdx = ExcelHelpers.FindColumnHeader(Row, "ANLG#");
                result = Math.Min(result, ANLGNumIdx); // Result will be -1 if column does not exist.
                ClientTagIdx = ExcelHelpers.FindColumnHeader(Row, "ClientTag");
                result = Math.Min(result, ClientTagIdx); // Result will be -1 if column does not exist.
                DescriptionIdx = ExcelHelpers.FindColumnHeader(Row, "Description");
                result = Math.Min(result, DescriptionIdx); // Result will be -1 if column does not exist.
                UnitIdx = ExcelHelpers.FindColumnHeader(Row, "Unit");
                result = Math.Min(result, UnitIdx); // Result will be -1 if column does not exist.
                MinIdx = ExcelHelpers.FindColumnHeader(Row, "Min.");
                result = Math.Min(result, MinIdx); // Result will be -1 if column does not exist.
                MaxIdx = ExcelHelpers.FindColumnHeader(Row, "Max.");
                result = Math.Min(result, MaxIdx); // Result will be -1 if column does not exist.
                SFLoLimitIdx = ExcelHelpers.FindColumnHeader(Row, "SFLoLimit");
                result = Math.Min(result, SFLoLimitIdx); // Result will be -1 if column does not exist.
                SFHiLimitIdx = ExcelHelpers.FindColumnHeader(Row, "SFHiLimit");
                result = Math.Min(result, SFHiLimitIdx); // Result will be -1 if column does not exist.
                LoAlmLimitIdx = ExcelHelpers.FindColumnHeader(Row, "LoAlmLimit");
                result = Math.Min(result, LoAlmLimitIdx); // Result will be -1 if column does not exist.
                LoSDLimitIdx = ExcelHelpers.FindColumnHeader(Row, "LoSDLimit");
                result = Math.Min(result, LoSDLimitIdx); // Result will be -1 if column does not exist.
                HiAlmLimitIdx = ExcelHelpers.FindColumnHeader(Row, "HiAlmLimit");
                result = Math.Min(result, HiAlmLimitIdx); // Result will be -1 if column does not exist.
                HiSDLimitIdx = ExcelHelpers.FindColumnHeader(Row, "HiSDLimit");
                result = Math.Min(result, HiSDLimitIdx); // Result will be -1 if column does not exist.
                OpenLimitIdx = ExcelHelpers.FindColumnHeader(Row, "OpenLimit");
                result = Math.Min(result, OpenLimitIdx); // Result will be -1 if column does not exist.
                ConversionFormulaIdx = ExcelHelpers.FindColumnHeader(Row, "ConversionFormula");
                result = Math.Min(result, ConversionFormulaIdx); // Result will be -1 if column does not exist.
                // TODO - Determine if Source of ANLG_IN is needed.
                //SourceIdx = ExcelHelpers.FindColumnHeader(Row, "Source");
                //result = Math.Min(result, SourceIdx); // Result will be -1 if column does not exist.
                return result > -1; // Result is -1 if any column does not exist.
            }

            public override string ToString()
            {
                return ANLGNum;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class AO_INSheet : ISheet
    {
        public AO_INSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        internal List<AO_INRow> Rows { get; set; } = new List<AO_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "AO_IN";
        public SorterToolImporter Parent { get; }

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
                var item = new AO_INRow(this);
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

                Rows = new List<AO_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        AO_INRow item = new AO_INRow(this);
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

        #region Classes

        public class AO_INRow
        {
            public AO_INRow(AO_INSheet aO_INSheet)
            {
                this.Parent = aO_INSheet;
            }

            #region Properties

            public string Address { get; set; }

            //TODO - This prob shouldn't be used.
            public string ClientTag { get; set; }

            public string Description { get; set; }

            public double? EngMax { get; set; }

            public double? EngMin { get; set; }

            public int? Index { get; set; }

            public string LocationOnNetwork { get; set; }

            /// <summary>
            /// Module point reference (i.e. AI01_CH01)
            /// </summary>
            public string ModulePoint { get; set; }

            public string SignalType { get; set; }

            public string SlotNumber { get; set; }

            public string Specifier { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["#"]) ? default(int?) : Convert.ToInt32(row["#"]);

                if (Index == default(int?) || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                LocationOnNetwork = Convert.IsDBNull(row["Location on Network"]) ? "" : Convert.ToString(row["Location on Network"]);
                SlotNumber = Convert.IsDBNull(row["Slot #"]) ? "" : Convert.ToString(row["Slot #"]);
                ModulePoint = Convert.IsDBNull(row["Module Point"]) ? "" : Convert.ToString(row["Module Point"]);
                ClientTag = Convert.IsDBNull(row["Client Tag"]) ? "" : Convert.ToString(row["Client Tag"]);
                SignalType = Convert.IsDBNull(row["Signal Type"]) ? "" : Convert.ToString(row["Signal Type"]);
                EngMin = Convert.IsDBNull(row["Eng Min"]) ? default(double?) : Convert.ToDouble(row["Eng Min"]);
                EngMax = Convert.IsDBNull(row["Eng Max"]) ? default(double?) : Convert.ToDouble(row["Eng Max"]);
                Address = Convert.IsDBNull(row["Address"]) ? "" : Convert.ToString(row["Address"]);
                Specifier = Convert.IsDBNull(row["SPECIFIER"]) ? "" : Convert.ToString(row["SPECIFIER"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    LocationOnNetwork = Row.Range["B1"].Value?.ToString().Trim() ?? "";
                    SlotNumber = Row.Range["C1"].Value?.ToString().Trim() ?? "";
                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["D1"]).Trim();
                    ClientTag = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    Description = Row.Range["H1"].Value?.ToString().Trim() ?? "";
                    SignalType = Row.Range["I1"].Value?.ToString().Trim() ?? "";
                    EngMin = ExcelHelpers.ExcelCellToDouble(Row.Range["K1"]);
                    EngMax = ExcelHelpers.ExcelCellToDouble(Row.Range["L1"]);
                    Address = Row.Range["M1"].Value?.ToString().Trim() ?? "";
                    Specifier = Row.Range["O1"].Value?.ToString().Trim() ?? "";

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
                return ModulePoint;
            }

            public I_O_LISTSheet.I_O_LISTRow I_O_List
            {
                get
                {
                    I_O_LISTSheet.I_O_LISTRow alarmIn = Parent.Parent.FindI_O_LISTRow(ModulePoint);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in I_O_LIST for {this}.");
                        return null;
                    }
                }
            }

            public AO_INSheet Parent { get; }

            #endregion Methods
        }

        #endregion Classes
    }

    public class Bently_INSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        internal List<Bently_InRow> Rows { get; set; } = new List<Bently_InRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "Bently_IN";

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
                var item = new Bently_InRow();
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

                Rows = new List<Bently_InRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        Bently_InRow item = new Bently_InRow();
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

        #region Classes

        public class Bently_InRow
        {
            #region Properties

            public int? AnlgNum { get; set; }

            public string Description { get; set; }

            public double? EngMax { get; set; }

            public double? EngMin { get; set; }

            public string GapAddress { get; set; }

            public string HH_SP_Addr { get; set; }

            public string H_SP_Addr { get; set; }

            public int? Index { get; set; }

            public string LL_SP_Addr { get; set; }

            public string L_SP_Addr { get; set; }

            public string ModulePoint { get; private set; }

            public string StatusAddress { get; set; }

            public string ValueAddress { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["#"]) ? default(int?) : Convert.ToInt32(row["#"]);

                if (Index == default(int?) || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                ModulePoint = Convert.IsDBNull(row["Module Point"]) ? "" : Convert.ToString(row["Module Point"]);
                ValueAddress = Convert.IsDBNull(row["VALUE_ADDR"]) ? "" : Convert.ToString(row["VALUE_ADDR"]);
                StatusAddress = Convert.IsDBNull(row["STAT_ADDR"]) ? "" : Convert.ToString(row["STAT_ADDR"]);
                GapAddress = Convert.IsDBNull(row["GAP_ADDR"]) ? "" : Convert.ToString(row["GAP_ADDR"]);
                LL_SP_Addr = Convert.IsDBNull(row["LL_SP_ADDR"]) ? "" : Convert.ToString(row["LL_SP_ADDR"]);
                L_SP_Addr = Convert.IsDBNull(row["L_SP_ADDR"]) ? "" : Convert.ToString(row["L_SP_ADDR"]);
                H_SP_Addr = Convert.IsDBNull(row["H_SP_ADDR"]) ? "" : Convert.ToString(row["H_SP_ADDR"]);
                HH_SP_Addr = Convert.IsDBNull(row["HH_SP_ADDR"]) ? "" : Convert.ToString(row["HH_SP_ADDR"]);
                AnlgNum = Convert.IsDBNull(row["ANLG#"]) ? default(int?) : Convert.ToInt32(row["ANLG#"]);
                EngMin = Convert.IsDBNull(row["Eng Min"]) ? default(double?) : Convert.ToDouble(row["Eng Min"]);
                EngMax = Convert.IsDBNull(row["Eng Max"]) ? default(double?) : Convert.ToDouble(row["Eng Max"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["D1"]).Trim();
                    AnlgNum = ExcelHelpers.ExcelCellToInt32(Row.Range["E1"]);
                    Description = Row.Range["H1"].Value?.ToString().Trim() ?? "";
                    EngMin = ExcelHelpers.ExcelCellToDouble(Row.Range["K1"]);
                    EngMax = ExcelHelpers.ExcelCellToDouble(Row.Range["L1"]);
                    ValueAddress = Row.Range["AA1"].Value?.ToString().Trim() ?? "";
                    GapAddress = Row.Range["AB1"].Value?.ToString().Trim() ?? "";
                    LL_SP_Addr = Row.Range["AC1"].Value?.ToString().Trim() ?? "";
                    L_SP_Addr = Row.Range["AD1"].Value?.ToString().Trim() ?? "";
                    H_SP_Addr = Row.Range["AE1"].Value?.ToString().Trim() ?? "";
                    HH_SP_Addr = Row.Range["AF1"].Value?.ToString().Trim() ?? "";
                    StatusAddress = Row.Range["AG1"].Value?.ToString().Trim() ?? "";

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
                return ModulePoint;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class BN_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private SorterToolImporter Parent;

        #endregion Fields

        #region Constructors

        public BN_MAPSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        #endregion Constructors

        #region Properties

        internal List<BN_MAPRow> Rows { get; set; } = new List<BN_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "BN_MAP";

        #endregion Properties

        #region Methods

        public bool CheckSheet()
        {
            bool sheetOkay = true;
            foreach (var row in Rows)
            {
                // Check for correct Modbus Address in the range of 45001-45500 or 46001-47000
                if ((row.BNAddress < 45001 || row.BNAddress > 45500) && (row.BNAddress < 46001 || row.BNAddress > 47000))
                {
                    Trace.TraceWarning($"In BN_MAP sheet, {row.ModulePoint} has a Modbus address {row.BNAddress} " +
                      $"outside the valid range of 45001-45500 or 46001-47000.");
                    sheetOkay = false;
                }
            }

            return sheetOkay;
        }

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
                var item = new BN_MAPRow(this);
                if (item.ImportRow(row))
                    Rows.Add(item);
            }

            // Perform check
            CheckSheet();

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

                Rows = new List<BN_MAPRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        BN_MAPRow item = new BN_MAPRow(this);
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                // Perform check
                CheckSheet();

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

        #region Classes

        public class BN_MAPRow
        {
            #region Fields

            private BN_MAPSheet Parent;

            #endregion Fields

            #region Constructors

            public BN_MAPRow(BN_MAPSheet bN_MAPSheet)
            {
                this.Parent = bN_MAPSheet;
            }

            #endregion Constructors

            #region Properties

            public I_O_LISTSheet.I_O_LISTRow I_O_List
            {
                get
                {
                    I_O_LISTSheet.I_O_LISTRow alarmIn = Parent.Parent.FindI_O_LISTRow(ModulePoint);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in I_O_LIST for {this}.");
                        return null;
                    }
                }
            }

            /// <summary>
            /// The ModulePoint reference (i.e. BN02_CH01). This is the 'Slot Reference' field in the Excel file.
            /// </summary>
            public string ModulePoint { get; set; }

            public string PLCAddress { get; set; }
            public int? BNAddress { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                ModulePoint = Convert.IsDBNull(row[(int)ExcelCol.A]) ? "" : Convert.ToString(row[(int)ExcelCol.A]);

                if (ModulePoint == "") return false;

                PLCAddress = Convert.IsDBNull(row[(int)ExcelCol.G]) ? "" : Convert.ToString(row[(int)ExcelCol.G]);
                BNAddress = Convert.IsDBNull(row[(int)ExcelCol.H]) ? default(int) : Convert.ToInt32(row[(int)ExcelCol.H]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["A1"]).Trim();

                    if (ModulePoint == "") return false;

                    PLCAddress = ExcelHelpers.ExcelCellToString(Row.Range["G1"]).Trim();
                    BNAddress = ExcelHelpers.ExcelCellToInt32(Row.Range["H1"]);

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
                return ModulePoint;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class CAESheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 12;

        #endregion Fields

        #region Properties

        public List<CAERow> Rows { get; set; } = new List<CAERow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "C&E";

        #endregion Properties

        #region Methods

        public bool CheckSheet()
        {
            bool sheetOkay = true;
            foreach (var row in Rows)
            {
                // Check for spaces
                if (row.AlarmClassA.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Alarm Class A column.");
                    sheetOkay = false;
                }
                if (row.AlarmClassB.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Alarm Class B column.");
                    sheetOkay = false;
                }
                if (row.AlarmClassC.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Alarm Class C column.");
                    sheetOkay = false;
                }
                if (row.ShutdownClassA.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Shutdown Class A column.");
                    sheetOkay = false;
                }
                if (row.ShutdownClassB.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Shutdown Class B column.");
                    sheetOkay = false;
                }
                if (row.ShutdownClassC.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Shutdown Class C column.");
                    sheetOkay = false;
                }
                if (row.Vented.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Vented Shutdown column.");
                    sheetOkay = false;
                }
                if (row.NonVented.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the NonVented Shutdown column.");
                    sheetOkay = false;
                }
                if (row.SigFailAlarm.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Signal Fail Alarm column.");
                    sheetOkay = false;
                }
                if (row.SigFailShutdown.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Signal Fail Shutdown column.");
                    sheetOkay = false;
                }
                if (row.ShutdownBypassable.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Shutdown Bypassable column.");
                    sheetOkay = false;
                }
                if (row.VotingGroup.Contains(" "))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has space in the Voting Group column.");
                    sheetOkay = false;
                }

                // Other checks
                if ((row.ShutdownClass != "") && !row.IsVented && !row.IsNonVented)
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} has a Shutdown Class (A/B/C), but is not Vented or Non-Vented.");
                    sheetOkay = false;
                }

                if (((row.SignalType == "AI") || (row.SignalType == "TC") || (row.SignalType == "RTD") || (row.SignalType == "FRQ"))
                    && ((row.ShutdownClass != "") && (row.SigFailShutdown == "") && (row.VotingGroup == "")))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} is an Analog with Shutdown Class (A/B/C), but no Signal Fail Shutdown.");
                    sheetOkay = false;
                }

                if (((row.SignalType == "AI") || (row.SignalType == "TC") || (row.SignalType == "RTD") || (row.SignalType == "FRQ"))
                    && ((row.VotingGroup != "") && (row.SigFailAlarm == "")))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} is an Analog with Voting Group. Votied analogs need to have a " +
                       $"Signal Fail Alarm.");
                    sheetOkay = false;
                }

                if (((row.SignalType == "AI") || (row.SignalType == "TC") || (row.SignalType == "RTD") || (row.SignalType == "FRQ"))
                    && ((row.VotingGroup != "") && (row.SigFailShutdown != "")))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} is an Analog with Voting Group. Voted analogs cannot have a " +
                       $"Signal Fail Shutdown.");
                    sheetOkay = false;
                }

                // TODO - This can be removed when it is not a requirment to manually generate LL/HH alarms for voted analogs.
                if (((row.SignalType == "AI") || (row.SignalType == "TC") || (row.SignalType == "RTD") || (row.SignalType == "FRQ"))
                && (row.VotingGroup != "") &&
                ((row.ShutdownClassA.Contains("LL") && !row.AlarmClassA.Contains("LL"))
                || (row.ShutdownClassA.Contains("HH") && !row.AlarmClassA.Contains("HH"))
                || (row.ShutdownClassB.Contains("LL") && !row.AlarmClassB.Contains("LL"))
                || (row.ShutdownClassB.Contains("HH") && !row.AlarmClassB.Contains("HH"))
                || (row.ShutdownClassC.Contains("LL") && !row.AlarmClassC.Contains("LL"))
                || (row.ShutdownClassC.Contains("HH") && !row.AlarmClassC.Contains("HH"))
                ))
                {
                    Trace.TraceWarning($"In C&E sheet, {row.ModulePoint} is an Analog with Voting Group. " +
                       $"HH and/or LL needs to be defined in the Alarm column as well as the Shutdown column.");
                    sheetOkay = false;
                }
            }

            return sheetOkay;
        }

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
                var item = new CAERow();
                if (item.ImportRow(row))
                    Rows.Add(item);
            }

            // Perform check
            CheckSheet();

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

                Rows = new List<CAERow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        CAERow item = new CAERow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                // Perform check
                CheckSheet();

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

        #region Classes

        public class CAERow
        {
            #region Properties

            /// <summary>
            /// Returns the Alarm Class ('A', 'B' or 'C')
            /// </summary>
            public string AlarmClass
            {
                get
                {
                    if (AlarmClassA != "")
                        return "A";
                    else if (AlarmClassB != "")
                        return "B";
                    else if (AlarmClassC != "")
                        return "C";
                    else
                        return "";
                }
            }

            /// <summary>
            /// Class A alarm (i.e. L/H)
            /// </summary>
            public string AlarmClassA { get; set; }

            /// <summary>
            /// Class B alarm (i.e. L/H)
            /// </summary>
            public string AlarmClassB { get; set; }

            /// <summary>
            /// Class C alarm (i.e. L/H)
            /// </summary>
            public string AlarmClassC { get; set; }

            public string Description { get; set; }

            /// <summary>
            /// Indicates that the shutdown is Non-Vented.
            /// </summary>
            public bool IsNonVented
            {
                get
                {
                    return NonVented.ToLower().Trim() == "x";
                }
            }

            /// <summary>
            /// Indicates that the shutdown is Vented.
            /// </summary>
            public bool IsVented
            {
                get
                {
                    return Vented.ToLower().Trim() == "x";
                }
            }

            /// <summary>
            /// Module point reference (i.e. AI01_CH01)
            /// </summary>
            public string ModulePoint { get; set; }

            /// <summary>
            /// Value of the Non-Vented column (i.e. empty string or 'X')
            /// </summary>
            public string NonVented { get; set; }

            /// <summary>
            /// Shutdown bypassable. Either blank or X.
            /// </summary>
            public dynamic ShutdownBypassable { get; private set; }

            /// <summary>
            /// Returns the Shutdown Class ('A', 'B' or 'C')
            /// </summary>
            public string ShutdownClass
            {
                get
                {
                    if (ShutdownClassA != "")
                        return "A";
                    else if (ShutdownClassB != "")
                        return "B";
                    else if (ShutdownClassC != "")
                        return "C";
                    else
                        return "";
                }
            }

            /// <summary>
            /// Class A shutdown (i.e. LL/HH/D)
            /// </summary>
            public string ShutdownClassA { get; set; }

            /// <summary>
            /// Class B shutdown (i.e. LL/HH/D)
            /// </summary>
            public string ShutdownClassB { get; set; }

            /// <summary>
            /// Class C shutdown (i.e. LL/HH/D)
            /// </summary>
            public string ShutdownClassC { get; set; }

            /// <summary>
            /// Signal Fail Alarm column checked. Either blank or X.
            /// </summary>
            public dynamic SigFailAlarm { get; private set; }

            /// <summary>
            /// Signal Fail Shutdown column checked. Either blank or X.
            /// </summary>
            public dynamic SigFailShutdown { get; private set; }

            /// <summary>
            /// Returns the signal type (i.e. DI, DO, AI, AO, TC, RTD, FRQ)
            /// </summary>
            public string SignalType { get; set; }

            /// <summary>
            /// Value of the Vented column (i.e. empty string or 'X')
            /// </summary>
            public string Vented { get; set; }

            /// <summary>
            /// Voting group (i.e. 'A', 'B', 'C')
            /// </summary>
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
                //TODO - Read Sorter Tool CAE using ExcelDataReader (needs more work as headers not on 1st row)
                //Index = Convert.IsDBNull(row["Number"]) ? default(int?) : Convert.ToInt32(row["Number"]);

                //if (Index == default(int?) || Index == 0) return false;

                //Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                //Tag = Convert.IsDBNull(row["Instrument Tag"]) ? "" : Convert.ToString(row["Instrument Tag"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["A1"]).Trim();

                    // Do not import spare lines
                    if (string.IsNullOrWhiteSpace(ModulePoint)) return false;

                    Description = ExcelHelpers.ExcelCellToString(Row.Range["E1"]).Trim();
                    SignalType = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    AlarmClassA = Row.Range["K1"].Value?.ToString().Trim() ?? "";
                    AlarmClassB = Row.Range["L1"].Value?.ToString().Trim() ?? "";
                    AlarmClassC = Row.Range["M1"].Value?.ToString().Trim() ?? "";
                    ShutdownClassA = Row.Range["N1"].Value?.ToString().Trim() ?? "";
                    ShutdownClassB = Row.Range["O1"].Value?.ToString().Trim() ?? "";
                    ShutdownClassC = Row.Range["P1"].Value?.ToString().Trim() ?? "";
                    Vented = Row.Range["S1"].Value?.ToString().Trim() ?? "";
                    NonVented = Row.Range["T1"].Value?.ToString().Trim() ?? "";
                    SigFailAlarm = Row.Range["W1"].Value?.ToString().Trim() ?? "";
                    SigFailShutdown = Row.Range["X1"].Value?.ToString().Trim() ?? "";
                    ShutdownBypassable = Row.Range["Y1"].Value?.ToString().Trim() ?? "";
                    VotingGroup = Row.Range["Z1"].Value?.ToString().Trim() ?? "";

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
                return ModulePoint;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class CONFC_INSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        public List<CONFC_INRow> Rows { get; set; } = new List<CONFC_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "CONFC_IN";

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
                var item = new CONFC_INRow();
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

                Rows = new List<CONFC_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        CONFC_INRow item = new CONFC_INRow();
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

        #region Classes

        public class CONFC_INRow
        {
            #region Properties

            public string Description { get; set; }

            public int? Index { get; set; }

            public string Tag { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["Number"]) ? default(int?) : Convert.ToInt32(row["Number"]);

                if (Index == default(int?) || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                Tag = Convert.IsDBNull(row["Instrument Tag"]) ? "" : Convert.ToString(row["Instrument Tag"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    Tag = ExcelHelpers.ExcelCellToString(Row.Range["B1"]).Trim();
                    Description = ExcelHelpers.ExcelCellToString(Row.Range["C1"]).Trim();

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
                return $"CONFC{Index:000}";
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class DI_INSheet : SorterToolSheet, ISheet
    {
        public DI_INSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        internal List<DI_INRow> Rows { get; set; } = new List<DI_INRow>();

        public override string WorksheetName => "DI_IN";

        public SorterToolImporter Parent { get; }

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
                var item = new DI_INRow(this);
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

                Rows = new List<DI_INRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        DI_INRow item = new DI_INRow(this);
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        DI_INRow item = new DI_INRow(this);
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
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

        #region Classes

        public class DI_INRow
        {
            public DI_INRow(DI_INSheet dI_INSheet)
            {
                this.Parent = dI_INSheet;
            }

            #region Properties

            public string Address { get; set; }

            public static int AddressIdx { get; private set; }

            public static int BitIdx { get; private set; }

            public string Description { get; set; }

            public static int DescriptionIdx { get; private set; }

            public int? Index { get; set; }

            public string LocationOnNetwork { get; set; }

            public static int LocationonNetworkIdx { get; private set; }

            /// <summary>
            /// Module point reference (i.e. AI01_CH01)
            /// </summary>
            public string ModulePoint { get; set; }

            public static int ModulePointIdx { get; private set; }

            public string SignalType { get; set; }

            public static int SignalTypeIdx { get; private set; }

            public string SlotNumber { get; set; }

            public static int SlotNumIdx { get; private set; }

            public string Specifier { get; set; }

            public static int SPECIFIERIdx { get; private set; }

            public string ClientTag { get; set; }

            public static int ClientTagIdx { get; private set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["Bit"]) ? default(int?) : Convert.ToInt32(row["Bit"]);

                if (Index == default(int?) || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                LocationOnNetwork = Convert.IsDBNull(row["Location on Network"]) ? "" : Convert.ToString(row["Location on Network"]);
                SlotNumber = Convert.IsDBNull(row["Slot #"]) ? "" : Convert.ToString(row["Slot #"]);
                ModulePoint = Convert.IsDBNull(row["Module Point"]) ? "" : Convert.ToString(row["Module Point"]);
                ClientTag = Convert.IsDBNull(row["Client Tag"]) ? "" : Convert.ToString(row["Client Tag"]);
                SignalType = Convert.IsDBNull(row["Signal Type"]) ? "" : Convert.ToString(row["Signal Type"]);
                Address = Convert.IsDBNull(row["Address"]) ? "" : Convert.ToString(row["Address"]);
                Specifier = Convert.IsDBNull(row["SPECIFIER"]) ? "" : Convert.ToString(row["SPECIFIER"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    //Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    //if (Index == default(int?) || Index == 0) return false;

                    //LocationOnNetwork = Row.Range["B1"].Value?.ToString().Trim() ?? "";
                    //SlotNumber = Row.Range["C1"].Value?.ToString().Trim() ?? "";
                    //ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["D1"]).Trim();
                    //Description = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    //SignalType = Row.Range["H1"].Value?.ToString().Trim() ?? "";
                    //Address = Row.Range["I1"].Value?.ToString().Trim() ?? "";
                    //Specifier = Row.Range["K1"].Value?.ToString().Trim() ?? "";

                    Index = ExcelHelpers.ExcelCellToInt32(Row.Columns[BitIdx]);

                    if (Index == default(int?) || Index == 0) return false;

                    LocationOnNetwork = ExcelHelpers.ExcelCellToString(Row.Columns[LocationonNetworkIdx]);
                    SlotNumber = ExcelHelpers.ExcelCellToString(Row.Columns[SlotNumIdx]);
                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Columns[ModulePointIdx]);
                    ClientTag = ExcelHelpers.ExcelCellToString(Row.Columns[ClientTagIdx]);
                    Description = ExcelHelpers.ExcelCellToString(Row.Columns[DescriptionIdx]);
                    SignalType = ExcelHelpers.ExcelCellToString(Row.Columns[SignalTypeIdx]);
                    Address = ExcelHelpers.ExcelCellToString(Row.Columns[AddressIdx]);
                    Specifier = ExcelHelpers.ExcelCellToString(Row.Columns[SPECIFIERIdx]);

                    return true;
                }
                catch (Exception ex)
                {
                    string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                    Trace.TraceError(errMsg);
                    throw new ApplicationException(errMsg);
                }
            }

            public bool ReadColumnHeaders(Range Row)
            {
                int result = 0;
                BitIdx = ExcelHelpers.FindColumnHeader(Row, "Bit");
                result = Math.Min(result, BitIdx); // Result will be -1 if column does not exist.
                LocationonNetworkIdx = ExcelHelpers.FindColumnHeader(Row, "Location On Network");
                result = Math.Min(result, LocationonNetworkIdx); // Result will be -1 if column does not exist.
                SlotNumIdx = ExcelHelpers.FindColumnHeader(Row, "Slot #");
                result = Math.Min(result, SlotNumIdx); // Result will be -1 if column does not exist.
                ModulePointIdx = ExcelHelpers.FindColumnHeader(Row, "Module Point");
                result = Math.Min(result, ModulePointIdx); // Result will be -1 if column does not exist.
                DescriptionIdx = ExcelHelpers.FindColumnHeader(Row, "Description");
                result = Math.Min(result, DescriptionIdx); // Result will be -1 if column does not exist.
                ClientTagIdx = ExcelHelpers.FindColumnHeader(Row, "Client Tag");
                result = Math.Min(result, ClientTagIdx); // Result will be -1 if column does not exist.
                SignalTypeIdx = ExcelHelpers.FindColumnHeader(Row, "Signal Type");
                result = Math.Min(result, SignalTypeIdx); // Result will be -1 if column does not exist.
                AddressIdx = ExcelHelpers.FindColumnHeader(Row, "Address");
                result = Math.Min(result, AddressIdx); // Result will be -1 if column does not exist.
                SPECIFIERIdx = ExcelHelpers.FindColumnHeader(Row, "SPECIFIER");
                result = Math.Min(result, SPECIFIERIdx); // Result will be -1 if column does not exist.
                return result > -1; // Result is -1 if any column does not exist.
            }

            public override string ToString()
            {
                return ModulePoint;
            }

            #endregion Methods

            public I_O_LISTSheet.I_O_LISTRow I_O_List
            {
                get
                {
                    I_O_LISTSheet.I_O_LISTRow alarmIn = Parent.Parent.FindI_O_LISTRow(ModulePoint);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in I_O_LIST for {this}.");
                        return null;
                    }
                }
            }

            public DI_INSheet Parent { get; }
        }

        #endregion Classes
    }

    public class DO_INSheet : ISheet
    {
        public DO_INSheet(SorterToolImporter sorterToolImporter)
        {
            this.Parent = sorterToolImporter;
        }

        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        internal List<DO_INRow> Rows { get; set; } = new List<DO_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "DO_IN";
        public SorterToolImporter Parent { get; }

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
                var item = new DO_INRow(this);
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

                Rows = new List<DO_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        DO_INRow item = new DO_INRow(this);
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

        #region Classes

        public class DO_INRow
        {
            #region Properties

            public DO_INRow(DO_INSheet dO_INSheet)
            {
                this.Parent = dO_INSheet;
            }

            public string Address { get; set; }

            public string Description { get; set; }

            public int? Index { get; set; }

            public string LocationOnNetwork { get; set; }

            /// <summary>
            /// Module point reference (i.e. AI01_CH01)
            /// </summary>
            public string ModulePoint { get; set; }

            public string SignalType { get; set; }

            public string SlotNumber { get; set; }

            public string Specifier { get; set; }

            public string ClientTag { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["Bit"]) ? default(int?) : Convert.ToInt32(row["Bit"]);

                if (Index == default(int?) || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                LocationOnNetwork = Convert.IsDBNull(row["Location on Network"]) ? "" : Convert.ToString(row["Location on Network"]);
                SlotNumber = Convert.IsDBNull(row["Slot #"]) ? "" : Convert.ToString(row["Slot #"]);
                ModulePoint = Convert.IsDBNull(row["Module Point"]) ? "" : Convert.ToString(row["Module Point"]);
                ClientTag = Convert.IsDBNull(row["Client Tag"]) ? "" : Convert.ToString(row["Client Tag"]);
                SignalType = Convert.IsDBNull(row["Signal Type"]) ? "" : Convert.ToString(row["Signal Type"]);
                Address = Convert.IsDBNull(row["Address"]) ? "" : Convert.ToString(row["Address"]);
                Specifier = Convert.IsDBNull(row["SPECIFIER"]) ? "" : Convert.ToString(row["SPECIFIER"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    LocationOnNetwork = Row.Range["B1"].Value?.ToString().Trim() ?? "";
                    SlotNumber = Row.Range["C1"].Value?.ToString().Trim() ?? "";
                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["D1"]).Trim();
                    ClientTag = Row.Range["F1"].Value?.ToString().Trim() ?? "";
                    Description = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    SignalType = Row.Range["H1"].Value?.ToString().Trim() ?? "";
                    Address = Row.Range["I1"].Value?.ToString().Trim() ?? "";
                    Specifier = Row.Range["K1"].Value?.ToString().Trim() ?? "";

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
                return ModulePoint;
            }

            #endregion Methods

            public I_O_LISTSheet.I_O_LISTRow I_O_List
            {
                get
                {
                    I_O_LISTSheet.I_O_LISTRow alarmIn = Parent.Parent.FindI_O_LISTRow(ModulePoint);
                    if (alarmIn != null)
                        return alarmIn;
                    else
                    {
                        Trace.TraceWarning($"During replacement in {Parent.WorksheetName}, could not find instance in I_O_LIST for {this}.");
                        return null;
                    }
                }
            }

            public DO_INSheet Parent { get; }
        }

        #endregion Classes
    }

    public class I_O_LISTSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        public List<I_O_LISTRow> Rows { get; set; } = new List<I_O_LISTRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "I_O_LIST";

        #endregion Properties

        #region Methods

        public bool CheckSheet()
        {
            List<string> usedDRTags = new List<string>();
            List<string> usedClientTags = new List<string>();

            bool sheetOkay = true;
            foreach (var row in Rows)
            {
                // Check that the BN channels are starting with CH01 instead of CH00 in order to match BN Modbus export.
                if (row.SignalType == "BN" && row.Channel == "00")
                {
                    Trace.TraceWarning($"In I_O_LIST sheet, {row.ModulePoint} has a channel number of '00'. " +
                        $"As this is a BN point, channels need to start with '01' to match BN Modbus export.");
                    sheetOkay = false;
                }

                // Check if DR Tagname already exists
                string dRTag = row.DRTagNo;
                if (!string.IsNullOrEmpty(dRTag))
                {
                    if (usedDRTags.Contains(dRTag))
                    {
                        Trace.TraceWarning($"In I_O_LIST sheet, {row.ModulePoint} has a DRTagNo of '{dRTag}' " +
                            $"which is already used. Should be unique.");
                        sheetOkay = false;
                    }
                    else
                    {
                        usedDRTags.Add(dRTag);
                    }
                }

                // Check if Client Tagname already exists
                string clientTag = row.ClientTagNo;
                if (!string.IsNullOrEmpty(clientTag))
                {
                    if (usedClientTags.Contains(clientTag))
                    {
                        Trace.TraceWarning($"In I_O_LIST sheet, {row.ModulePoint} has a ClientTagNo of '{clientTag}' " +
                            $"which is already used. Should be unique.");
                        sheetOkay = false;
                    }
                    else
                    {
                        usedClientTags.Add(clientTag);
                    }
                }
            }

            return sheetOkay;
        }

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
                var item = new I_O_LISTRow();
                if (item.ImportRow(row))
                    Rows.Add(item);
            }

            // Perform check
            CheckSheet();

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

                Rows = new List<I_O_LISTRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        I_O_LISTRow item = new I_O_LISTRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                // Perform check
                CheckSheet();

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

        #region Classes

        public class I_O_LISTRow
        {
            #region Fields

            private string _channel;
            private string _modulePoint;
            private string _rack;
            private string _rackSlot;
            private string _slot;

            #endregion Fields

            #region Properties

            /// <summary>
            /// The bit (Status) for discrete IO or Anlg number for analog IO of the IO point (i.e. 1 to n)
            /// </summary>
            public int? BitANLG { get; set; }

            /// <summary>
            /// The Channel in the I/O List, usually 2 digit (i.e. 00, 01 ... 16)
            /// </summary>
            public string Channel
            {
                get { return _channel; }
            }

            public string ClientTagNo { get; private set; }

            public string DRTagNo { get; private set; }

            public double? EngHigh { get; private set; }

            public double? EngLow { get; private set; }

            public string EngUnits { get; private set; }

            public double? H { get; private set; }

            public double? HH { get; private set; }

            /// <summary>
            /// Module Point returned in the safe format DI01_CH01 (no dashes or slashes)
            /// </summary>
            public string ModulePointSafe { get; private set; }

            public string RawDescription { get; private set; }

            public double? L { get; private set; }

            public double? LL { get; private set; }

            /// <summary>
            /// The Module Point in one of the following formats: DI-01/03, DI-01_03 or DI01_CH03
            /// </summary>
            public string ModulePoint
            {
                get { return _modulePoint; }
                private set
                {
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        // tt[t]-[r]ss/cc (DI-01/03), tt[t]-[r]ss_cc (DI-01_03) or tt[t][r]ss_CHcc (DI01_CH03)
                        if (value.Contains("/"))
                        {
                            // tt[t]-[r]ss/cc (DI-01/03)
                            string[] splitStr = ExcelHelpers.SplitString(value);

                            string type = value.Substring(0, value.IndexOf("-", StringComparison.Ordinal));
                            _rackSlot = value.Substring(value.IndexOf("-", StringComparison.Ordinal) + 1,
                                value.IndexOf("/", StringComparison.Ordinal) - value.IndexOf("-", StringComparison.Ordinal) - 1);
                            if (_rackSlot.Length == 2)
                            {
                                _rack = "0";
                                _slot = _rackSlot;
                            }
                            else if (_rackSlot.Length == 3)
                            {
                                _rack = _rackSlot.Substring(0, 1);
                                _slot = _rackSlot.Substring(1, 2);
                            }
                            _channel = value.Substring(value.IndexOf("/", StringComparison.Ordinal) + 1);

                            ModulePointSafe = $"{type}{_rackSlot}_CH{_channel}";
                        }
                        else if (value.Contains("-"))
                        {
                            // tt[t]-[r]ss_cc (DI-01_03)
                            string type = value.Substring(0, value.IndexOf("-", StringComparison.Ordinal));
                            _rackSlot = value.Substring(value.IndexOf("-", StringComparison.Ordinal) + 1,
                                value.IndexOf("_", StringComparison.Ordinal) - value.IndexOf("-", StringComparison.Ordinal) - 1);
                            if (_rackSlot.Length == 2)
                            {
                                _rack = "0";
                                _slot = _rackSlot;
                            }
                            else if (_rackSlot.Length == 3)
                            {
                                _rack = _rackSlot.Substring(0, 1);
                                _slot = _rackSlot.Substring(1, 2);
                            }
                            _channel = value.Substring(value.IndexOf("_", StringComparison.Ordinal) + 1);

                            ModulePointSafe = $"{type}{_rackSlot}_CH{_channel}";
                        }
                        else if (value.Contains("_"))
                        {
                            // tt[t][r]ss[x]_CHcc (DI01_CH03) or

                            string type = "";
                            _rackSlot = "0";
                            _rack = "0";
                            _slot = "0";
                            _channel = "0";

                            string[] all = value.Split('_');

                            string[] ttss = ExcelHelpers.SplitString(all[0]);
                            string[] cc = ExcelHelpers.SplitString(all[1]);

                            if (ttss.Length == 2) // DI|01
                            {
                                type = ttss[0];
                                _rackSlot = ttss[1];

                                if (ttss[1].Length == 2)
                                {
                                    _rack = "0";
                                    _slot = ttss[1];
                                }
                                else if (ttss[1].Length == 3)
                                {
                                    _rack = ttss[1].Substring(0, 1);
                                    _slot = ttss[1].Substring(1, 2);
                                }
                            }
                            else if (ttss.Length == 3) // DI|01|A
                            {
                                type = ttss[0];

                                _rackSlot = ttss[1] + ttss[2];

                                if (ttss[1].Length == 2)
                                {
                                    _rack = "0";
                                    _slot = ttss[1] + ttss[2];
                                }
                                else if (ttss[1].Length == 3)
                                {
                                    _rack = ttss[1].Substring(0, 1);
                                    _slot = ttss[1].Substring(1, 2) + ttss[2];
                                }
                            }
                            if (cc.Length == 2)
                            {
                                _channel = cc[1];
                            }

                            ModulePointSafe = $"{type}{_rackSlot}_CH{_channel}";
                        }
                        else
                        {
                            Trace.TraceWarning($"Unknown format of Module Point in IO List for '{value}'. Does not match format of DI-01/03, DI-01_03 or DI01_CH03.");
                        }
                    }
                    _modulePoint = value;
                }
            }

            public string Notes { get; private set; }

            public double? OP { get; private set; }

            public string Rack
            {
                get { return _rack; }
            }

            public string RackSlot
            {
                get { return _rackSlot; }
            }

            /// <summary>
            /// Returns the signal type (i.e. DI, DO, AI, AO, TC, RTD, FRQ)
            /// </summary>
            public string SignalType { get; private set; }

            public string Slot
            {
                get { return _slot; }
            }

            #endregion Properties

            #region Methods

            public string Description(bool TagBased)
            {
                if (TagBased)
                    return $"{RawDescription} {Tagname(!TagBased)} [{BitANLG}]";
                else
                    return $"{Tagname(!TagBased)} / {RawDescription} [{BitANLG}]";
            }

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                try
                {
                    ModulePoint = Convert.IsDBNull(row["Module /  Point"]) ? "" : Convert.ToString(row["Module /  Point"]);

                    if (string.IsNullOrWhiteSpace(ModulePoint))
                        return false;

                    DRTagNo = Convert.IsDBNull(row["D-R Tag No."]) ? "" : Convert.ToString(row["D-R Tag No."]);
                    BitANLG = Convert.IsDBNull(row["Bit/ANLG"]) ? default(int?) : Convert.ToInt32(row["Bit/ANLG"]);
                    ClientTagNo = Convert.IsDBNull(row["ClientTagNo."]) ? "" : Convert.ToString(row["ClientTagNo."]);
                    RawDescription = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                    SignalType = Convert.IsDBNull(row["Signal Type"]) ? "" : Convert.ToString(row["Signal Type"]);
                    EngUnits = Convert.IsDBNull(row["Units / Dev Type"]) ? "" : Convert.ToString(row["Units / Dev Type"]);
                    EngLow = Convert.IsDBNull(row["Eng Low"]) ? default(double?) : Convert.ToDouble(row["Eng Low"]);
                    EngHigh = Convert.IsDBNull(row["Eng High"]) ? default(double?) : Convert.ToDouble(row["Eng High"]);
                    Notes = Convert.IsDBNull(row["Notes"]) ? "" : Convert.ToString(row["Notes"]);
                    L = Convert.IsDBNull(row["L"]) ? default(double?) : Convert.ToDouble(row["L"]);
                    LL = Convert.IsDBNull(row["LL"]) ? default(double?) : Convert.ToDouble(row["LL"]);
                    H = Convert.IsDBNull(row["H"]) ? default(double?) : Convert.ToDouble(row["H"]);
                    HH = Convert.IsDBNull(row["HH"]) ? default(double?) : Convert.ToDouble(row["HH"]);
                    OP = Convert.IsDBNull(row["OP"]) ? default(double?) : Convert.ToDouble(row["OP"]);

                    return true;
                }
                catch (Exception ex)
                {
                    Trace.TraceError(ex.Message);
                    throw new System.InvalidOperationException($"Error importing row: {ex.Message}");
                }
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    ModulePoint = ExcelHelpers.ExcelCellToString(Row.Range["B1"]);

                    if (string.IsNullOrWhiteSpace(ModulePoint))
                        return false;

                    DRTagNo = ExcelHelpers.ExcelCellToString(Row.Range["C1"]).Trim();
                    BitANLG = ExcelHelpers.ExcelCellToInt32(Row.Range["D1"]);
                    ClientTagNo = Row.Range["E1"].Value?.ToString().Trim() ?? "";
                    RawDescription = Row.Range["F1"].Value?.ToString().Trim() ?? "";
                    SignalType = Row.Range["G1"].Value?.ToString().Trim() ?? "";
                    EngUnits = Row.Range["H1"].Value?.ToString().Trim() ?? "";
                    EngLow = ExcelHelpers.ExcelCellToDouble(Row.Range["I1"]);
                    EngHigh = ExcelHelpers.ExcelCellToDouble(Row.Range["J1"]);
                    Notes = Row.Range["K1"].Value?.ToString().Trim() ?? "";
                    L = ExcelHelpers.ExcelCellToDouble(Row.Range["L1"]);
                    LL = ExcelHelpers.ExcelCellToDouble(Row.Range["M1"]);
                    H = ExcelHelpers.ExcelCellToDouble(Row.Range["N1"]);
                    HH = ExcelHelpers.ExcelCellToDouble(Row.Range["O1"]);
                    OP = ExcelHelpers.ExcelCellToDouble(Row.Range["P1"]);

                    return true;
                }
                catch (Exception ex)
                {
                    string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                    Trace.TraceError(errMsg);
                    throw new ApplicationException(errMsg);
                }
            }

            public string TagnameSafe(bool TagBased)
            {
                return Tagname(TagBased).Replace("-", "_");
            }

            public string DRTagNoSafe
            {
                get
                {
                    return DRTagNo.Replace("-", "_");
                }
            }

            public string ClientTagNoSafe
            {
                get
                {
                    return ClientTagNo.Replace("-", "_");
                }
            }

            /// <summary>
            /// Gets the Tagname of the IO Point
            /// </summary>
            /// <param name="TagBased">Selects tag or IO address based tag.</param>
            /// <returns>Returns the client tag if tagbased selected, otherwise the IO address.</returns>
            public string Tagname(bool TagBased)
            {
                if (TagBased)
                    if (!string.IsNullOrWhiteSpace(ClientTagNo))
                        return ClientTagNo;
                    else
                        return ModulePointSafe;
                else
                    return ModulePointSafe;
            }

            public override string ToString()
            {
                return ModulePoint;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    /// <summary>
    /// Modbus Map row, common to all Modbus sheets
    /// </summary>
    public class MB_MAPRow
    {
        #region Properties

        public string Access { get; private set; }

        public static int AccessIdx { get; private set; }

        public string Alarm { get; private set; }

        public static int AlarmIdx { get; private set; }

        public int? AlmNum { get; private set; }

        public static int AlmNumIdx { get; private set; }

        public string Analog { get; private set; }

        public static int AnalogIdx { get; private set; }

        public int? AnlgNum { get; private set; }

        public static int AnlgNumIdx { get; private set; }

        public string Bit { get; private set; }

        public string BitAddr { get; private set; }

        public static int BitAddrIdx { get; private set; }

        public static int BitIdx { get; private set; }

        public int? SIEDBNum { get; private set; }

        public static int SIEDBNumIdx { get; private set; }

        public string PLCTagname { get; private set; }

        public static int PLCTagnameIdx { get; private set; }

        public string DBSubType { get; private set; }

        public static int DBSubTypeIdx { get; private set; }

        public string Description { get; private set; }

        public static int DescriptionIdx { get; private set; }

        public string ClientDescription { get; private set; }

        public static int ClientDescriptionIdx { get; private set; }

        public string ClientTag { get; private set; }

        public static int ClientTagIdx { get; private set; }

        public double? EngMax { get; private set; }

        public static int EngMaxIdx { get; private set; }

        public double? EngMin { get; private set; }

        public static int EngMinIdx { get; private set; }

        public string FCSubType { get; private set; }

        public static int FCSubTypeIdx { get; private set; }

        public string FType { get; private set; }

        public static int FTypeIdx { get; private set; }

        public int? Index { get; private set; }

        //SortedDictionary<string, int> columns = new SortedDictionary<string, int>();

        public static int IndexIdx { get; private set; }

        public string IndexNum { get; private set; }

        public static int IndexNumIdx { get; private set; }

        public string IndexNumPlus1 { get; private set; }

        public static int IndexNumPlus1Idx { get; private set; }

        public string MBAddress { get; private set; }

        public static int MBAddressIdx { get; private set; }

        public string Name { get; private set; }

        public static int NameIdx { get; private set; }

        public string OFFSET { get; private set; }

        public static int OFFSETIdx { get; private set; }

        public string DBName { get; private set; }

        public static int DBNameIdx { get; private set; }

        public string PLCAddr { get; private set; }

        public string PLCAddr2 { get; private set; }

        public static int PLCAddr2Idx { get; private set; }

        public static int PLCAddrIdx { get; private set; }

        public double? RawMax { get; private set; }

        public static int RawMaxIdx { get; private set; }

        public double? RawMin { get; private set; }

        public static int RawMinIdx { get; private set; }

        public string RegisterScaling { get; private set; }

        public static int RegisterScalingIdx { get; private set; }

        public string SD { get; private set; }

        public static int SDIdx { get; private set; }

        public int? StatNum { get; private set; }

        public static int StatNumIdx { get; private set; }

        public string Status { get; private set; }

        public static int StatusIdx { get; private set; }

        public string Units { get; private set; }

        public static int UnitsIdx { get; private set; }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Creator for import tag
        /// </summary>
        /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
        /// <returns>True is the row can be imported, else false.</returns>
        public bool ImportRow(DataRow row)
        {
            // TODO Implement reading of DataRow?
            //Index = Convert.IsDBNull(row[(int)ExcelCol.A]) ? default(int?) : Convert.ToInt32(row[(int)ExcelCol.A]);

            //if (Index == default(int?) || Index == 0) return false;

            //AlarmRef = Convert.IsDBNull(row[(int)ExcelCol.B]) ? "" : Convert.ToString(row[(int)ExcelCol.B]);
            //AlarmDescription = Convert.IsDBNull(row["Alarm Description"]) ? "" : Convert.ToString(row["Alarm Description"]);
            //AlarmNumber = Convert.IsDBNull(row["Alarm Number"]) ? default(int?) : Convert.ToInt32(row["Alarm Number"]);
            //AlarmTag = Convert.IsDBNull(row["Alarm Tag"]) ? "" : Convert.ToString(row["Alarm Tag"]);
            //AlarmType = Convert.IsDBNull(row["Alarm Type"]) ? "" : Convert.ToString(row["Alarm Type"]);

            return true;
        }

        public bool ImportRow(Range Row)
        {
            try
            {
                Index = ExcelHelpers.ExcelCellToInt32(Row.Columns[IndexIdx]);

                if (Index == default || Index == 0) return false;

                FCSubType = ExcelHelpers.ExcelCellToString(Row.Columns[FCSubTypeIdx]);
                DBSubType = ExcelHelpers.ExcelCellToString(Row.Columns[DBSubTypeIdx]);
                DBName = ExcelHelpers.ExcelCellToString(Row.Columns[DBNameIdx]);
                PLCAddr = ExcelHelpers.ExcelCellToString(Row.Columns[PLCAddrIdx]);
                PLCAddr2 = ExcelHelpers.ExcelCellToString(Row.Columns[PLCAddr2Idx]);
                BitAddr = ExcelHelpers.ExcelCellToString(Row.Columns[BitAddrIdx]);
                AnlgNum = ExcelHelpers.ExcelCellToInt32(Row.Columns[AnlgNumIdx]);
                StatNum = ExcelHelpers.ExcelCellToInt32(Row.Columns[StatNumIdx]);
                AlmNum = ExcelHelpers.ExcelCellToInt32(Row.Columns[AlmNumIdx]);
                IndexNum = ExcelHelpers.ExcelCellToString(Row.Columns[IndexNumIdx]);
                IndexNumPlus1 = ExcelHelpers.ExcelCellToString(Row.Columns[IndexNumPlus1Idx]);
                PLCTagname = ExcelHelpers.ExcelCellToString(Row.Columns[PLCTagnameIdx]);
                SIEDBNum = ExcelHelpers.ExcelCellToInt32(Row.Columns[SIEDBNumIdx]);
                RawMin = ExcelHelpers.ExcelCellToDouble(Row.Columns[RawMinIdx]);
                RawMax = ExcelHelpers.ExcelCellToDouble(Row.Columns[RawMaxIdx]);
                Analog = ExcelHelpers.ExcelCellToString(Row.Columns[AnalogIdx]);
                Status = ExcelHelpers.ExcelCellToString(Row.Columns[StatusIdx]);
                Alarm = ExcelHelpers.ExcelCellToString(Row.Columns[AlarmIdx]);
                SD = ExcelHelpers.ExcelCellToString(Row.Columns[SDIdx]);
                MBAddress = ExcelHelpers.ExcelCellToString(Row.Columns[MBAddressIdx]);
                Bit = ExcelHelpers.ExcelCellToString(Row.Columns[BitIdx]);
                FType = ExcelHelpers.ExcelCellToString(Row.Columns[FTypeIdx]);
                Access = ExcelHelpers.ExcelCellToString(Row.Columns[AccessIdx]);
                Name = ExcelHelpers.ExcelCellToString(Row.Columns[NameIdx]);
                Description = ExcelHelpers.ExcelCellToString(Row.Columns[DescriptionIdx]);
                ClientTag = ExcelHelpers.ExcelCellToString(Row.Columns[ClientTagIdx]);
                ClientDescription = ExcelHelpers.ExcelCellToString(Row.Columns[ClientDescriptionIdx]);
                Units = ExcelHelpers.ExcelCellToString(Row.Columns[UnitsIdx]);
                EngMin = ExcelHelpers.ExcelCellToDouble(Row.Columns[EngMinIdx]);
                EngMax = ExcelHelpers.ExcelCellToDouble(Row.Columns[EngMaxIdx]);
                OFFSET = ExcelHelpers.ExcelCellToString(Row.Columns[OFFSETIdx]);
                RegisterScaling = ExcelHelpers.ExcelCellToString(Row.Columns[RegisterScalingIdx]);

                return true;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel row {Row.Row} in {Row.Worksheet.Name}: {ex.Message}";
                Trace.TraceError(errMsg);
                throw new ApplicationException(errMsg);
            }
        }

        public bool ReadColumnHeaders(Range Row)
        {
            //columns = new SortedDictionary<string, int>();
            //string colName = "";
            //for (int i = 1; i <= Row.Columns.Count; i++)
            //{
            //    colName = Row.Columns[i].Value?.ToString() ?? "";
            //    if (colName == "")
            //        colName = $"Column{ExcelHelpers.ColumnLetter(i, false)}";

            //    columns.Add(colName, i);
            //}

            int result = 0;
            IndexIdx = ExcelHelpers.FindColumnHeader(Row, "Index");
            result = Math.Min(result, IndexIdx); // Result will be -1 if column does not exist.
            FCSubTypeIdx = ExcelHelpers.FindColumnHeader(Row, "FCSubType");
            result = Math.Min(result, FCSubTypeIdx); // Result will be -1 if column does not exist.
            DBSubTypeIdx = ExcelHelpers.FindColumnHeader(Row, "DBSubType");
            result = Math.Min(result, DBSubTypeIdx); // Result will be -1 if column does not exist.
            DBNameIdx = ExcelHelpers.FindColumnHeader(Row, "DBName");
            result = Math.Min(result, DBNameIdx); // Result will be -1 if column does not exist.
            PLCAddrIdx = ExcelHelpers.FindColumnHeader(Row, "PLCAddr");
            result = Math.Min(result, PLCAddrIdx); // Result will be -1 if column does not exist.
            PLCAddr2Idx = ExcelHelpers.FindColumnHeader(Row, "PLCAddr2");
            result = Math.Min(result, PLCAddr2Idx); // Result will be -1 if column does not exist.
            BitAddrIdx = ExcelHelpers.FindColumnHeader(Row, "BitAddr");
            result = Math.Min(result, BitAddrIdx); // Result will be -1 if column does not exist.
            AnlgNumIdx = ExcelHelpers.FindColumnHeader(Row, "AnlgNum");
            result = Math.Min(result, AnlgNumIdx); // Result will be -1 if column does not exist.
            StatNumIdx = ExcelHelpers.FindColumnHeader(Row, "StatNum");
            result = Math.Min(result, StatNumIdx); // Result will be -1 if column does not exist.
            AlmNumIdx = ExcelHelpers.FindColumnHeader(Row, "AlmNum");
            result = Math.Min(result, AlmNumIdx); // Result will be -1 if column does not exist.
            SIEDBNumIdx = ExcelHelpers.FindColumnHeader(Row, "SIE DBNum");
            result = Math.Min(result, SIEDBNumIdx); // Result will be -1 if column does not exist.
            PLCTagnameIdx = ExcelHelpers.FindColumnHeader(Row, "PLC Tagname");
            result = Math.Min(result, PLCTagnameIdx); // Result will be -1 if column does not exist.
            IndexNumIdx = ExcelHelpers.FindColumnHeader(Row, "IndexNum");
            result = Math.Min(result, IndexNumIdx); // Result will be -1 if column does not exist.
            IndexNumPlus1Idx = ExcelHelpers.FindColumnHeader(Row, "IndexNumPlus1");
            result = Math.Min(result, IndexNumPlus1Idx); // Result will be -1 if column does not exist.
            RawMinIdx = ExcelHelpers.FindColumnHeader(Row, "Raw Min");
            result = Math.Min(result, RawMinIdx); // Result will be -1 if column does not exist.
            RawMaxIdx = ExcelHelpers.FindColumnHeader(Row, "Raw Max");
            result = Math.Min(result, RawMaxIdx); // Result will be -1 if column does not exist.
            AnalogIdx = ExcelHelpers.FindColumnHeader(Row, "Analog");
            result = Math.Min(result, AnalogIdx); // Result will be -1 if column does not exist.
            StatusIdx = ExcelHelpers.FindColumnHeader(Row, "Status");
            result = Math.Min(result, StatusIdx); // Result will be -1 if column does not exist.
            AlarmIdx = ExcelHelpers.FindColumnHeader(Row, "Alarm");
            result = Math.Min(result, AlarmIdx); // Result will be -1 if column does not exist.
            SDIdx = ExcelHelpers.FindColumnHeader(Row, "SD");
            result = Math.Min(result, SDIdx); // Result will be -1 if column does not exist.
            MBAddressIdx = ExcelHelpers.FindColumnHeader(Row, "MB_Address");
            result = Math.Min(result, MBAddressIdx); // Result will be -1 if column does not exist.
            BitIdx = ExcelHelpers.FindColumnHeader(Row, "Bit");
            result = Math.Min(result, BitIdx); // Result will be -1 if column does not exist.
            FTypeIdx = ExcelHelpers.FindColumnHeader(Row, "FType");
            result = Math.Min(result, FTypeIdx); // Result will be -1 if column does not exist.
            AccessIdx = ExcelHelpers.FindColumnHeader(Row, "Access");
            result = Math.Min(result, AccessIdx); // Result will be -1 if column does not exist.
            NameIdx = ExcelHelpers.FindColumnHeader(Row, "Name");
            result = Math.Min(result, NameIdx); // Result will be -1 if column does not exist.
            DescriptionIdx = ExcelHelpers.FindColumnHeader(Row, "Description");
            result = Math.Min(result, DescriptionIdx); // Result will be -1 if column does not exist.
            ClientTagIdx = ExcelHelpers.FindColumnHeader(Row, "Client Tag");
            result = Math.Min(result, ClientTagIdx); // Result will be -1 if column does not exist.
            ClientDescriptionIdx = ExcelHelpers.FindColumnHeader(Row, "Client Description");
            result = Math.Min(result, ClientDescriptionIdx); // Result will be -1 if column does not exist.
            UnitsIdx = ExcelHelpers.FindColumnHeader(Row, "Units");
            result = Math.Min(result, UnitsIdx); // Result will be -1 if column does not exist.
            EngMinIdx = ExcelHelpers.FindColumnHeader(Row, "EngMin");
            result = Math.Min(result, EngMinIdx); // Result will be -1 if column does not exist.
            EngMaxIdx = ExcelHelpers.FindColumnHeader(Row, "EngMax");
            result = Math.Min(result, EngMaxIdx); // Result will be -1 if column does not exist.
            OFFSETIdx = ExcelHelpers.FindColumnHeader(Row, "OFFSET");
            result = Math.Min(result, OFFSETIdx); // Result will be -1 if column does not exist.
            RegisterScalingIdx = ExcelHelpers.FindColumnHeader(Row, "Register Scaling ");
            result = Math.Min(result, RegisterScalingIdx); // Result will be -1 if column does not exist.

            return result > -1; // Result is -1 if any column does not exist.
        }

        public override string ToString()
        {
            return $"Index{Index:000}";
        }

        #endregion Methods
    }

    public class MB1_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB1_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
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

    public class MB2_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB2_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class MB3_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB3_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class MB4_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB4_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class MB5_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB5_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class MB6_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB6_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class MB7_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB7_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class MB8_MAPSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;
        private readonly long headerRow = 1;

        #endregion Fields

        #region Properties

        public List<MB_MAPRow> Rows { get; set; } = new List<MB_MAPRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "MB8_MAP";

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
                var item = new MB_MAPRow();
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
                return false;
                //throw new ApplicationException(errMsg);
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
                return false;
                //throw new ApplicationException(errMsg);
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

                Rows = new List<MB_MAPRow>();

                // Read in Rows to Objects
                bool headersOkay = false;
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row == headerRow)
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        // TODO - Apply this to all Excel reads
                        headersOkay = item.ReadColumnHeaders(inputRow);
                    }
                    if (headersOkay && (inputRow.Row >= firstDataRow))
                    {
                        MB_MAPRow item = new MB_MAPRow();
                        if (item.ImportRow(inputRow))
                            Rows.Add(item);
                    }
                }

                return headersOkay;
            }
            catch (Exception ex)
            {
                string errMsg = $"Error importing Excel sheet '{WorksheetName}': {ex.Message}";
                Trace.TraceError(errMsg);
                return false;
                //throw new ApplicationException(errMsg);
            }
        }

        public override string ToString()
        {
            return WorksheetName;
        }

        #endregion Methods
    }

    public class SD_INSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        public List<SD_INRow> Rows { get; set; } = new List<SD_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "SD_IN";

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
                var item = new SD_INRow();
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

                Rows = new List<SD_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        SD_INRow item = new SD_INRow();
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

        #region Classes

        public class SD_INRow
        {
            #region Properties

            public string AlarmDescription { get; set; }

            /// <summary>
            /// Alarm number of the shutdown, typically starting at 513
            /// </summary>
            public int? AlarmNumber { get; set; }

            public string AlarmRef { get; set; }

            public string AlarmTag { get; set; }

            /// <summary>
            /// Alarm type (i.e. L, LL, HH_Voted)
            /// </summary>
            public string AlarmType { get; set; }
            public string CustomPLCTag { get; set; }
            public int? Index { get; set; }

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

                AlarmRef = Convert.IsDBNull(row[(int)ExcelCol.B]) ? "" : Convert.ToString(row[(int)ExcelCol.B]);
                AlarmDescription = Convert.IsDBNull(row["Alarm Description"]) ? "" : Convert.ToString(row["Alarm Description"]);
                AlarmNumber = Convert.IsDBNull(row["Alarm Number"]) ? default(int?) : Convert.ToInt32(row["Alarm Number"]);
                AlarmTag = Convert.IsDBNull(row["Alarm Tag"]) ? "" : Convert.ToString(row["Alarm Tag"]);
                AlarmType = Convert.IsDBNull(row["Alarm Type"]) ? "" : Convert.ToString(row["Alarm Type"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    AlarmRef = ExcelHelpers.ExcelCellToString(Row.Range["B1"]).Trim();
                    AlarmDescription = Row.Range["C1"].Value?.ToString().Trim() ?? "";

                    // In AB Sorter Tool, T1 is used instead of G1
                    // TODO - replace with the platform selection in the HOME sheet
                    if ((Row.Range["T1"].Value?.ToString() ?? "") != "")
                        AlarmNumber = ExcelHelpers.ExcelCellToInt32(Row.Range["T1"]);
                    else
                        AlarmNumber = ExcelHelpers.ExcelCellToInt32(Row.Range["G1"]);

                    AlarmTag = Row.Range["I1"].Value?.ToString().Trim() ?? "";
                    AlarmType = Row.Range["K1"].Value?.ToString().Trim() ?? "";

                    CustomPLCTag = ExcelHelpers.ExcelCellToString(Row.Range["U1"]).Trim();
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
                return AlarmRef;
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class STAT_INSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        public List<STAT_INRow> Rows { get; set; } = new List<STAT_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "STAT_IN";

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
                var item = new STAT_INRow();
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

                Rows = new List<STAT_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        STAT_INRow item = new STAT_INRow();
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

        #region Classes

        public class STAT_INRow
        {
            #region Properties

            public string Description { get; set; }

            public int? Index { get; set; }

            public string OffMessage { get; private set; }

            public string OnMessage { get; private set; }

            public string PLCTag { get; set; }

            public string StatusRef { get; private set; }

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

                StatusRef = Convert.IsDBNull(row[(int)ExcelCol.B]) ? "" : Convert.ToString(row[(int)ExcelCol.B]);
                Description = Convert.IsDBNull(row[(int)ExcelCol.C]) ? "" : Convert.ToString(row[(int)ExcelCol.C]);
                PLCTag = Convert.IsDBNull(row[(int)ExcelCol.K]) ? "" : Convert.ToString(row[(int)ExcelCol.K]);
                OffMessage = Convert.IsDBNull(row[(int)ExcelCol.N]) ? "" : Convert.ToString(row[(int)ExcelCol.N]);
                OnMessage = Convert.IsDBNull(row[(int)ExcelCol.M]) ? "" : Convert.ToString(row[(int)ExcelCol.M]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    // TODO - The header of this sheet is all messed up. Needs to be fixed to read the data correctly.
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    StatusRef = ExcelHelpers.ExcelCellToString(Row.Range["B1"]).Trim();
                    Description = ExcelHelpers.ExcelCellToString(Row.Range["C1"]).Trim();
                    PLCTag = Row.Range["K1"].Value?.ToString().Trim() ?? "";
                    OnMessage = Row.Range["M1"].Value?.ToString().Trim() ?? "";
                    OffMessage = Row.Range["N1"].Value?.ToString().Trim() ?? "";

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
                return $"STAT{Index:000}";
            }

            #endregion Methods
        }

        #endregion Classes
    }

    public class TIMERS_INSheet : ISheet
    {
        #region Fields

        private readonly long firstDataRow = 2;

        #endregion Fields

        #region Properties

        internal List<TIMERS_INRow> Rows { get; set; } = new List<TIMERS_INRow>();

        public Workbook Workbook { get; set; }

        public Worksheet Worksheet { get; private set; }

        public string WorksheetName { get; private set; } = "TIMERS_IN";

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
                var item = new TIMERS_INRow();
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

                Rows = new List<TIMERS_INRow>();

                // Read in Rows to Objects
                foreach (Range inputRow in Worksheet.UsedRange.Rows)
                {
                    if (inputRow.Row >= firstDataRow)
                    {
                        TIMERS_INRow item = new TIMERS_INRow();
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

        #region Classes

        public class TIMERS_INRow
        {
            #region Properties

            public string Description { get; set; }

            public int? Index { get; set; }

            public int? Preset { get; set; }

            #endregion Properties

            #region Methods

            /// <summary>
            /// Creator for import tag
            /// </summary>
            /// <param name="row">A DataRow representing a line from a TriStation tag export file.</param>
            /// <returns>True is the row can be imported, else false.</returns>
            public bool ImportRow(DataRow row)
            {
                Index = Convert.IsDBNull(row["Number"]) ? default(int?) : Convert.ToInt32(row["Number"]);

                if (Index == default(int?) || Index == 0) return false;

                Description = Convert.IsDBNull(row["Description"]) ? "" : Convert.ToString(row["Description"]);
                Preset = Convert.IsDBNull(row["Preset in Seconds"]) ? default(int?) : Convert.ToInt32(row["Preset in Seconds"]);

                return true;
            }

            public bool ImportRow(Range Row)
            {
                try
                {
                    // TODO - The header of this sheet is all messed up. Needs to be fixed to read the data correctly.
                    Index = ExcelHelpers.ExcelCellToInt32(Row.Range["A1"]);

                    if (Index == default(int?) || Index == 0) return false;

                    Description = Row.Range["B1"].Value?.ToString().Trim() ?? "";
                    Preset = ExcelHelpers.ExcelCellToInt32(Row.Range["C1"]);

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
                return $"TIMER{Index:000}";
            }

            #endregion Methods
        }

        #endregion Classes
    }
}