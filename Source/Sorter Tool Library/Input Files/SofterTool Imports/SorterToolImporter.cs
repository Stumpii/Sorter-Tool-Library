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
    public class SorterToolImporter
    {
        #region Fields

        private Workbook _workbook;

        #endregion Fields

        #region Constructors

        public SorterToolImporter() : this(null)
        {
        }

        // TODO - Is this constructor really needs as this can use workbook or datareader?
        public SorterToolImporter(Workbook Workbook)
        {
            // TDOD - Add the other sheets
            AI_INSheet = new AI_INSheet(this);
            ALM_GENSheet = new ALM_GENSheet(this);
            AO_INSheet = new AO_INSheet(this);
            BN_MAPSheet = new BN_MAPSheet(this);
            DI_INSheet = new DI_INSheet(this);
            DO_INSheet = new DO_INSheet(this);
            SD_GENSheet = new SD_GENSheet(this);

            this.Workbook = Workbook;
        }

        #endregion Constructors

        #region Properties

        public AI_INSheet AI_INSheet { get; set; }

        public ALM_GENSheet ALM_GENSheet { get; set; }

        public ALM_INSheet ALM_INSheet { get; set; } = new ALM_INSheet();

        public ANLG_INSheet ANLG_INSheet { get; set; } = new ANLG_INSheet();

        public AO_INSheet AO_INSheet { get; set; }

        public Bently_INSheet Bently_INSheet { get; set; } = new Bently_INSheet();

        public BN_MAPSheet BN_MAPSheet { get; set; }

        public CAESheet CAESheet { get; set; } = new CAESheet();

        public CONFC_INSheet CONFC_INSheet { get; set; } = new CONFC_INSheet();

        public DI_INSheet DI_INSheet { get; set; }

        public DO_INSheet DO_INSheet { get; set; }

        public HOMESheet HOMESheet { get; set; } = new HOMESheet();

        public I_O_LISTSheet I_O_LISTSheet { get; set; } = new I_O_LISTSheet();

        public MB1_MAPSheet MB1_MAPSheet { get; set; } = new MB1_MAPSheet();
        public MB2_MAPSheet MB2_MAPSheet { get; set; } = new MB2_MAPSheet();
        public MB3_MAPSheet MB3_MAPSheet { get; set; } = new MB3_MAPSheet();
        public MB4_MAPSheet MB4_MAPSheet { get; set; } = new MB4_MAPSheet();
        public MB5_MAPSheet MB5_MAPSheet { get; set; } = new MB5_MAPSheet();
        public MB6_MAPSheet MB6_MAPSheet { get; set; } = new MB6_MAPSheet();
        public MB7_MAPSheet MB7_MAPSheet { get; set; } = new MB7_MAPSheet();
        public MB8_MAPSheet MB8_MAPSheet { get; set; } = new MB8_MAPSheet();

        public SD_GENSheet SD_GENSheet { get; set; }

        public SD_INSheet SD_INSheet { get; set; } = new SD_INSheet();

        public STAT_INSheet STAT_INSheet { get; set; } = new STAT_INSheet();

        public STD_DEV_SEL_3Sheet STD_DEV_SEL_3Sheet { get; set; } = new STD_DEV_SEL_3Sheet();

        public TIMERS_INSheet TIMERS_INSheet { get; set; } = new TIMERS_INSheet();

        public Workbook Workbook

        {
            get { return _workbook; }
            set
            {
                _workbook = value;
                AI_INSheet.Workbook = _workbook;
                ALM_GENSheet.Workbook = _workbook;
                ALM_INSheet.Workbook = _workbook;
                ANLG_INSheet.Workbook = _workbook;
                AO_INSheet.Workbook = _workbook;
                Bently_INSheet.Workbook = _workbook;
                BN_MAPSheet.Workbook = _workbook;
                CAESheet.Workbook = _workbook;
                CONFC_INSheet.Workbook = _workbook;
                DI_INSheet.Workbook = _workbook;
                DO_INSheet.Workbook = _workbook;
                HOMESheet.Workbook = _workbook;
                I_O_LISTSheet.Workbook = _workbook;
                MB1_MAPSheet.Workbook = _workbook;
                MB2_MAPSheet.Workbook = _workbook;
                MB3_MAPSheet.Workbook = _workbook;
                MB4_MAPSheet.Workbook = _workbook;
                MB5_MAPSheet.Workbook = _workbook;
                MB6_MAPSheet.Workbook = _workbook;
                MB7_MAPSheet.Workbook = _workbook;
                MB8_MAPSheet.Workbook = _workbook;
                SD_GENSheet.Workbook = _workbook;
                SD_INSheet.Workbook = _workbook;
                STAT_INSheet.Workbook = _workbook;
                STD_DEV_SEL_3Sheet.Workbook = _workbook;
                TIMERS_INSheet.Workbook = _workbook;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Find the row in the ALM_IN sheet with the matching alarm number.
        /// </summary>
        /// <param name="AlarmNo">The alarm number (i.e. 1 to 1024)</param>
        /// <returns>The matching row from the ALM_IN sheet.</returns>
        public ALM_INSheet.ALM_INRow FindALM_INRow(int? AlarmNo)
        {
            if (AlarmNo is null)
            {
                Trace.TraceWarning($"FindALM_INRow called with AlarmNo being null.");
                return null;
            }
            if (AlarmNo < 1)
            {
                Trace.TraceWarning($"FindALM_INRow called with AlarmNo being zero.");
                return null;
            }

            foreach (var item in ALM_INSheet.Rows)
            {
                if (item.AlarmNumber == AlarmNo)
                    return item;
            }

            Trace.TraceWarning($"FindALM_INRow called with AlarmNo {AlarmNo:000} which could not be found.");
            return null;
        }

        /// <summary>
        /// Find the row in the SD_IN sheet with the matching alarm number.
        /// </summary>
        /// <param name="AlarmNo">The alarm number (i.e. 1 to 1024)</param>
        /// <returns>The matching row from the SD_IN sheet.</returns>
        public SD_INSheet.SD_INRow FindSD_INRow(int? AlarmNo)
        {
            if (AlarmNo is null)
            {
                Trace.TraceWarning($"FindSD_INRow called with AlarmNo being null.");
                return null;
            }
            if (AlarmNo < 1)
            {
                Trace.TraceWarning($"FindSD_INRow called with AlarmNo being zero.");
                return null;
            }

            foreach (var item in SD_INSheet.Rows)
            {
                if (item.AlarmNumber == AlarmNo)
                    return item;
            }

            Trace.TraceWarning($"FindSD_INRow called with AlarmNo {AlarmNo:000} which could not be found.");
            return null;
        }

        /// <summary>
        /// Find the row in the STAT_IN sheet with the matching status number.
        /// </summary>
        /// <param name="StatNo">The status number (i.e. 1 to 1024)</param>
        /// <returns>The matching row from the STAT_IN sheet.</returns>
        public STAT_INSheet.STAT_INRow FindSTAT_INRow(int? StatNo)
        {
            if (StatNo is null)
            {
                Trace.TraceWarning($"FindSTAT_INRow called with StatNo being null.");
                return null;
            }
            if (StatNo < 1)
            {
                Trace.TraceWarning($"FindSTAT_INRow called with StatNo being zero.");
                return null;
            }

            foreach (var item in STAT_INSheet.Rows)
            {
                if (item.Index == StatNo)
                    return item;
            }

            Trace.TraceWarning($"FindSTAT_INRow called with StatNo {StatNo:000} which could not be found.");
            return null;
        }

        /// <summary>
        /// Find the row in the I_O_LIST sheet with the matching module point (i.e. AI01_CH01).
        /// </summary>
        /// <param name="ModulePoint">The module point (i.e. AI01_CH01)</param>
        /// <returns>The matching row from the I_O_LIST sheet.</returns>
        public I_O_LISTSheet.I_O_LISTRow FindI_O_LISTRow(string ModulePoint)
        {
            if (string.IsNullOrWhiteSpace(ModulePoint))
            {
                Trace.TraceWarning($"FindI_O_LISTRow called with ModulePoint blank.");
                return null;
            }

            foreach (var item in I_O_LISTSheet.Rows)
            {
                if (item.ModulePoint == ModulePoint)
                    return item;
            }

            Trace.TraceWarning($"FindI_O_LISTRow called with ModulePoint {ModulePoint} which could not be found.");
            return null;
        }

        public bool ReadSorterToolDataReader(DataSet DataSet)
        {
            AI_INSheet.ReadDataTable(DataSet);
            ALM_GENSheet.ReadDataTable(DataSet);
            ALM_INSheet.ReadDataTable(DataSet);
            ANLG_INSheet.ReadDataTable(DataSet);
            AO_INSheet.ReadDataTable(DataSet);
            Bently_INSheet.ReadDataTable(DataSet);
            BN_MAPSheet.ReadDataTable(DataSet);
            CAESheet.ReadDataTable(DataSet);
            CONFC_INSheet.ReadDataTable(DataSet);
            DI_INSheet.ReadDataTable(DataSet);
            DO_INSheet.ReadDataTable(DataSet);
            //HOMESheet.ReadDataTable(DataSet);
            I_O_LISTSheet.ReadDataTable(DataSet);
            MB1_MAPSheet.ReadDataTable(DataSet);
            MB2_MAPSheet.ReadDataTable(DataSet);
            MB3_MAPSheet.ReadDataTable(DataSet);
            MB4_MAPSheet.ReadDataTable(DataSet);
            MB5_MAPSheet.ReadDataTable(DataSet);
            MB6_MAPSheet.ReadDataTable(DataSet);
            MB7_MAPSheet.ReadDataTable(DataSet);
            MB8_MAPSheet.ReadDataTable(DataSet);
            SD_GENSheet.ReadDataTable(DataSet);
            SD_INSheet.ReadDataTable(DataSet);
            STAT_INSheet.ReadDataTable(DataSet);
            STD_DEV_SEL_3Sheet.ReadDataTable(DataSet);
            TIMERS_INSheet.ReadDataTable(DataSet);
            return true;
        }

        public bool ReadSorterToolDataReader(string path)
        {
            Trace.TraceInformation($"Opening file: {path}");
            Trace.Indent();

            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

            //Choose one of either 1 or 2
            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            //4. DataSet - Create column names from first row
            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();

            #region Replace table

            ///* Often an Excel sheet contains a table of data that does not start on the first row.
            // * There may be title, revision or notes rows before the table. The following code
            // * parses a table until the user determines that the real data begins (with or without
            // * column eaders). The subsequent rows are used to populate a temporary table. Once
            // * complete, the temporary table replaces the original table in the DataSet. */
            //string tablename = "ALM&SD"; // Define table name to be processed
            //bool isFirstRowAsColumnNames = true; // Configure if the first row of the real table contains column names

            //DataTable currentTable = result.Tables[tablename];
            //DataTable newTable = new DataTable(result.Tables[tablename].TableName);
            //bool foundTable = false;
            //long rowIndex = 0;
            //foreach (DataRow row in currentTable.Rows)
            //{
            //    // Check if this is the start of the real table
            //    if (!foundTable)
            //    {
            //        /* Place rules to determine if this is the start of the table.
            //         * Can be based on column value (row[i].ToSring() or based on row
            //         * index (rowIndex == n). */
            //        if ((row[2].ToString() == "Inputs / Outputs"))
            //        {
            //            // This is the start of the table
            //            if (isFirstRowAsColumnNames)
            //            {
            //                for (int i = 0; i < currentTable.Columns.Count; i++)
            //                {
            //                    if (row[i] != null && row[i].ToString().Length > 0)
            //                        AddColumnHandleDuplicate(newTable, row[i].ToString());
            //                    else
            //                        AddColumnHandleDuplicate(newTable, string.Concat("Column ", i + 1));
            //                }
            //            }
            //            else
            //            {
            //                // Create dummy column names
            //                for (int i = 0; i < currentTable.Columns.Count; i++)
            //                    newTable.Columns.Add(null, typeof(Object));

            //                // Import first row as data
            //                newTable.ImportRow(row);
            //            }
            //            foundTable = true;
            //        }
            //    }
            //    else
            //    {
            //        // Import data
            //        newTable.ImportRow(row);
            //    }
            //    rowIndex++;
            //}

            //// Replace old table with new table
            //result.Tables.Remove(currentTable);
            //result.Tables.Add(newTable);

            #endregion Replace table

            //// =========== TEST CODE TO PRINT COLUMN TYPES ===========
            //var tableNames = new List<string>() { "STD_DEV_3", "STD_DEV_3" };
            //foreach (var tbl in tableNames)
            //{
            //    DataColumnCollection colColl = result.Tables[tbl].Columns;
            //    Debug.WriteLine(string.Empty);
            //    Debug.WriteLine($"public class {tbl}\r\n{{");
            //    for (int i = 0; i < colColl.Count; i++)
            //    {
            //        Debug.WriteLine($"public {colColl[i].DataType.ToString().Replace("System.", "")}" +
            //            $" {colColl[i]}  {{ get; set; }}");
            //    }
            //    Debug.WriteLine("}");
            //    Debug.WriteLine(string.Empty);
            //}
            //// =========== END TEST CODE TO PRINT COLUMN TYPES ===========

            ReadSorterToolDataReader(result);

            Trace.Unindent();
            Trace.TraceInformation($"Finished reading file.");

            return true;
        }

        public bool ReadSorterToolExcel(Workbook Workbook)
        {
            AI_INSheet.ReadSheet(Workbook);
            ALM_GENSheet.ReadSheet(Workbook);
            ALM_INSheet.ReadSheet(Workbook);
            ANLG_INSheet.ReadSheet(Workbook);
            AO_INSheet.ReadSheet(Workbook);
            Bently_INSheet.ReadSheet(Workbook);
            BN_MAPSheet.ReadSheet(Workbook);
            CAESheet.ReadSheet(Workbook);
            CONFC_INSheet.ReadSheet(Workbook);
            DI_INSheet.ReadSheet(Workbook);
            DO_INSheet.ReadSheet(Workbook);
            HOMESheet.ReadSheet(Workbook);
            I_O_LISTSheet.ReadSheet(Workbook);
            MB1_MAPSheet.ReadSheet(Workbook);
            MB2_MAPSheet.ReadSheet(Workbook);
            MB3_MAPSheet.ReadSheet(Workbook);
            MB4_MAPSheet.ReadSheet(Workbook);
            MB5_MAPSheet.ReadSheet(Workbook);
            MB6_MAPSheet.ReadSheet(Workbook);
            MB7_MAPSheet.ReadSheet(Workbook);
            MB8_MAPSheet.ReadSheet(Workbook);
            SD_GENSheet.ReadSheet(Workbook);
            SD_INSheet.ReadSheet(Workbook);
            STAT_INSheet.ReadSheet(Workbook);
            STD_DEV_SEL_3Sheet.ReadSheet(Workbook);
            TIMERS_INSheet.ReadSheet(Workbook);
            return true;
        }

        #endregion Methods
    }
}