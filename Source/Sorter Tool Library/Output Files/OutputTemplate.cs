using ExcelDataReader;
using PropertyChanged;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace DRHMIConverter
{
    /// <summary>
    /// A wrapper for an Excel output template (i.e. WinCC 7.x Template.xlsx)
    /// </summary>
    [AddINotifyPropertyChangedInterface]
    public class OutputTemplate
    {
        #region "Fields"

        private List<OutputSheet> _outputSheets = new List<OutputSheet>();

        #endregion "Fields"

        #region "Properties"

        public string CommentText { get; private set; } = "COMMENT";
        public string SettingText { get; private set; } = "SETTING";
        public string FooterText { get; private set; } = "FOOTER";

        public string GroupByInputBeginText { get; private set; } = "GROUPBY_INPUT";

        public string GroupByOutputBeginText { get; private set; } = "GROUPBY_OUTPUT";

        public string GroupBySingletonBeginText { get; private set; } = "GROUPBY_SINGLE";

        public string GroupEndText { get; private set; } = "END_GROUP";

        public string HeaderText { get; private set; } = "HEADER";

        public List<OutputSheet> OutputSheets
        {
            get { return _outputSheets; }
            set { _outputSheets = value; }
        }

        #endregion "Properties"

        #region "Methods"

        /// <summary>
        /// Read in an xlsx template
        /// </summary>
        /// <param name="path"></param>
        public void ReadXLSXTemplate(string path)
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
            // TODO - enable UseHeaderRow. This means changing the following code to read the column header instead of row[0]
            DataSet result = excelReader.AsDataSet();
            //DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            //{
            //    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
            //    {
            //        UseHeaderRow = true
            //    }
            //});

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();

            #region Replace table

            ///* Often an Excel sheet contains a table of data that does not start on the first row.
            // * There may be title, revision or notes rows before the table. The following code
            // * parses a table until the user determines that the real data begins (with or without
            // * column eaders). The subsequent rows are used to populate a temporary table. Once
            // * complete, the temporary table replaces the original table in the DataSet. */
            //string tablename = "StructtypeElement"; // Define table name to be processed
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
            //        if ((subtypeColData == "Name"))
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

            OutputSheets.Clear();

            Trace.TraceInformation($"Parsing tables...");
            Trace.Indent();

            // Parse each worksheet (aa table)
            foreach (DataTable item in result.Tables)
            {
                Trace.TraceInformation($"Parsing {item.TableName}");

                // Read table/sheet
                OutputSheet outputSheet = new OutputSheet();
                outputSheet.SheetName = item.TableName;

                // Create default group
                TemplateGroup currentGroup = new TemplateGroup(outputSheet)
                {
                    groupBy = TemplateGroup.GroupBy.Output
                };
                outputSheet.MainGroup = currentGroup;

                // Skip sheet if first cell is not correct
                if (item.Rows[0][0].ToString() != "TYPE" &&
                    item.Rows[0][1].ToString() != "SUBTYPE" &&
                    item.Rows[0][2].ToString() != "RULE" &&
                    item.Rows[0][3].ToString() != "DATA")
                    continue;

                foreach (DataRow row in item.Rows)
                {
                    string typeColData = row[0].ToString();
                    string subtypeColData = row[1].ToString();
                    string ruleColData = row[2].ToString();
                    string dateColData = row[3].ToString();

                    // Skip header row
                    if (typeColData == "TYPE" &&
                        subtypeColData == "SUBTYPE" &&
                        ruleColData == "RULE" &&
                        dateColData == "DATA")
                        continue;

                    // Skip comment rows
                    if (typeColData == CommentText)
                        continue;

                    if (typeColData == SettingText)
                    {
                        string settingType = subtypeColData;
                        string settingValue = ruleColData;

                        switch (settingType.ToUpper())
                        {
                            case "FILENAME:":
                                Trace.TraceInformation($"Setting detected: {settingType}={settingValue}");
                                outputSheet.SheetName = settingValue;
                                break;

                            case "IGNORE SHEET:":
                                Trace.TraceInformation($"Setting detected: {settingType}={settingValue}");
                                outputSheet.IgnoreSheet = settingValue.ToUpper().Equals("TRUE");
                                break;

                            default:
                                Trace.TraceWarning($"Unknown setting detected: {settingType}={settingValue}");
                                break;
                        }
                    }
                    else if (typeColData == HeaderText)
                    {
                        OutputLine outputHeader = new OutputLine
                        {
                            // Read type from first column
                            Type = typeColData,
                            SubType = subtypeColData,
                            Rule = ruleColData
                        };

                        List<string> headerItems = new List<string>();
                        for (int i = 3; i < row.Table.Columns.Count; i++) // Skip first column
                        {
                            if (string.IsNullOrEmpty(row[i].ToString()))
                                break;

                            headerItems.Add(row[i].ToString());
                        }

                        // Add the header items
                        outputHeader.Items.AddRange(headerItems);

                        // Add the header
                        currentGroup.Headers.Add(outputHeader);
                    }
                    else if (typeColData == FooterText)
                    {
                        OutputLine outputFooter = new OutputLine
                        {
                            // Read type from first column
                            Type = typeColData,
                            SubType = subtypeColData,
                            Rule = ruleColData
                        };

                        List<string> footerItems = new List<string>();
                        for (int i = 3; i < row.Table.Columns.Count; i++) // Skip first column
                        {
                            if (string.IsNullOrEmpty(row[i].ToString()))
                                break;

                            footerItems.Add(row[i].ToString());
                        }

                        // Add the footer items
                        outputFooter.Items.AddRange(footerItems);

                        // Add the footer
                        currentGroup.Footers.Add(outputFooter);
                    }
                    else if (typeColData == GroupByInputBeginText)
                    {
                        // Start of new group (and end of last group)
                        // TODO - IF already in a group, why not make a sub-group?
                        TemplateGroup newGroup = new TemplateGroup(outputSheet);
                        newGroup.groupBy = TemplateGroup.GroupBy.Input;
                        newGroup.Parent = currentGroup;
                        currentGroup.Groups.Add(newGroup);
                        currentGroup = newGroup;
                    }
                    else if (typeColData == GroupByOutputBeginText)
                    {
                        // Start of new group (and end of last group)
                        // TODO - IF already in a group, why not make a sub-group?
                        TemplateGroup newGroup = new TemplateGroup(outputSheet);
                        newGroup.groupBy = TemplateGroup.GroupBy.Output;
                        newGroup.Parent = currentGroup;
                        currentGroup.Groups.Add(newGroup);
                        currentGroup = newGroup;
                    }
                    else if (typeColData == GroupBySingletonBeginText)
                    {
                        // Start of new group (and end of last group)
                        // TODO - IF already in a group, why not make a sub-group?
                        TemplateGroup newGroup = new TemplateGroup(outputSheet);
                        newGroup.groupBy = TemplateGroup.GroupBy.Singleton;
                        newGroup.Parent = currentGroup;
                        currentGroup.Groups.Add(newGroup);
                        currentGroup = newGroup;
                    }
                    else if (typeColData == GroupEndText)
                    {
                        // End of current group, use previous group
                        if (currentGroup.Parent != null)
                        {
                            currentGroup = currentGroup.Parent;
                        }
                        else
                        {
                            // Some error occurred
                        }
                    }
                    else // Must be data
                    {
                        OutputLine outputData = new OutputLine
                        {
                            // Read type from first column
                            Type = typeColData,
                            SubType = subtypeColData,
                            Rule = ruleColData
                        };

                        List<string> dataItem = new List<string>();
                        for (int i = 3; i < row.Table.Columns.Count; i++) // Skip first column
                        {
                            dataItem.Add(row[i].ToString());
                        }

                        // Add the data items
                        outputData.Items.AddRange(dataItem);

                        // Add the data
                        currentGroup.Data.Add(outputData);
                    }
                }

                Trace.TraceInformation($"Read {item.Rows.Count} rows.");

                OutputSheets.Add(outputSheet);
            }

            Trace.Unindent();
            Trace.TraceInformation($"Finished parsing tables...");
            Trace.Unindent();
            Trace.TraceInformation($"Finished with file: {path}");

            // return true;
        }

        #endregion "Methods"

        #region "Classes"

        /// <summary>
        /// A group of data (header, data and footer) in the output template.
        /// </summary>
        [AddINotifyPropertyChangedInterface]
        public class TemplateGroup
        {
            #region "Fields"

            public List<OutputLine> Data = new List<OutputLine>();
            public List<OutputLine> Footers = new List<OutputLine>();
            public List<TemplateGroup> Groups = new List<TemplateGroup>();
            public List<OutputLine> Headers = new List<OutputLine>();
            public TemplateGroup Parent { get; set; }

            #endregion "Fields"

            #region "Constructors"

            public TemplateGroup(OutputSheet outputSheet)
            {
                ParentSheet = outputSheet;
            }

            #endregion "Constructors"

            #region "Enums"

            public enum GroupBy
            {
                Input,
                Output,
                Singleton
            }

            #endregion "Enums"

            #region "Properties"

            public GroupBy groupBy { get; set; } = TemplateGroup.GroupBy.Output;

            public OutputSheet ParentSheet { get; }

            #endregion "Properties"
        }

        /// <summary>
        /// A row of data from the Excel sheet
        /// </summary>
        [AddINotifyPropertyChangedInterface]
        public class OutputLine
        {
            #region "Fields"

            public List<string> Items = new List<string>();

            #endregion "Fields"

            #region "Properties"

            public string SubType { get; set; }

            public string Type { get; set; }
            public string Rule { get; set; }

            #endregion "Properties"
        }

        /// <summary>
        /// An output sheet (a worksheet within the Excel template)
        /// </summary>
        [AddINotifyPropertyChangedInterface]
        public class OutputSheet
        {
            #region "Fields"

            public TemplateGroup MainGroup;

            #endregion "Fields"

            #region "Properties"

            public string SheetName { get; set; }
            public bool IgnoreSheet { get; internal set; }

            #endregion "Properties"
        }

        #endregion "Classes"
    }
}