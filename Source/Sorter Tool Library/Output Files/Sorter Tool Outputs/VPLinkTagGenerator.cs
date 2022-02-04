using DRHMIConverter;
using SorterToolLibrary;
using SorterToolLibrary.OutputBase;
using SorterToolLibrary.SorterTool;
using static DRHMIConverter.OutputTemplate;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using SorterToolLibrary.Output_Files.Sorter_Tool_Outputs;
using System.Diagnostics;

namespace VPLinkTagGenerator
{
    public abstract class OutputBase2
    {
        #region Fields

        internal string defaultOutputFile;

        //internal DRHMI_Config dRHMI_Config;
        internal OutputTemplate outputTemplate = new OutputTemplate();

        /// <summary>
        /// The output types supported by this template (i.e. ALM_GEN. STATUS, TIMER).
        /// </summary>
        internal List<IOutputType> outputTypes;

        internal HelperSettings programSettings;
        internal SorterToolImporter sorterTool;

        #endregion Fields

        #region Properties

        /// <summary>
        /// Defines the default output file extension (i.e. '.txt', '.awl', '.ini') when writing separate files.
        /// </summary>
        public string DefaultFileExtension { get; set; } = ".txt";

        public Encoding Encoding { get; set; } = Encoding.ASCII;

        /// <summary>
        /// For single file writes, makes a time stamped copy of the file.
        /// For separate file writes, makes a time stamped copy of the folder.
        /// </summary>
        public bool WriteTimeStampedCopy { get; set; } = true;

        public char Separator { get; set; } = ',';

        public string TemplateFilename { get; set; }

        /// <summary>
        /// Defines if outputing to a single file or multiple files.
        /// If True, it is assumed that the output file string represents a folder.
        /// </summary>
        public bool WriteSeparateFiles { get; set; } = false;

        #endregion Properties

        #region Methods

        /// <summary>
        /// Process an output group (list of lines in template)
        /// </summary>
        /// <param name="templateGroup">A group in an output template sheet (contains header
        /// rows, data rows, footer rows and possibly more groups).</param>
        /// <param name="outputSheet"></param>
        /// <returns></returns>
        public virtual StringBuilder ProcessGroup(TemplateGroup templateGroup, OutputSheet outputSheet)
        {
            StringBuilder tempSB = new StringBuilder();

            // Generate Header
            foreach (OutputLine item in templateGroup.Headers)
            {
                // Join items
                tempSB.AppendLine(string.Join(Separator.ToString(), item.Items));
            }

            // Generate Data
            switch (templateGroup.groupBy)
            {
                case TemplateGroup.GroupBy.Input:
                    // Group by Input

                    // Generate Data
                    foreach (IOutputType outputType in outputTypes)
                    {
                        StringBuilder tempStringBuilder = outputType.WriteInputData(templateGroup, Separator.ToString());
                        tempSB.Append(tempStringBuilder.ToString());
                    }
                    break;

                case TemplateGroup.GroupBy.Singleton:
                    // Generate a single instance for the group

                    // Process all lines
                    foreach (OutputLine line in templateGroup.Data)
                    {
                        // Define sheets that only require processing once
                        bool WriteSingleton = true;

                        // Generate Data
                        foreach (var outputType in outputTypes)
                        {
                            // Check if this is the right output type (i.e. ANLG)
                            // TODO - Why is this check outside the WriteOutputData, but inside the WriteInputData?
                            if (outputType.Type == line.Type)
                            {
                                string outputLine = string.Join(Separator.ToString(), line.Items);
                                StringBuilder tempStringBuilder = outputType.WriteOutputData(outputSheet.SheetName,
                                   outputLine, line, WriteSingleton);
                                tempSB.Append(tempStringBuilder.ToString());
                            }
                        }
                    }
                    break;

                case TemplateGroup.GroupBy.Output:
                    // Group by Output

                    // Process all lines in this template group
                    foreach (OutputLine line in templateGroup.Data)
                    {
                        // Define sheets that only require processing once
                        bool WriteSingleton = false;

                        // Parse each output template type
                        foreach (var outputType in outputTypes)
                        {
                            // Check if this is the right output type (i.e. ANLG)
                            if (outputType.Type == line.Type)
                            {
                                string outputLine = string.Join(Separator.ToString(), line.Items);
                                StringBuilder tempStringBuilder = outputType.WriteOutputData(outputSheet.SheetName,
                                   outputLine, line, WriteSingleton);

                                tempSB.Append(tempStringBuilder.ToString());
                            }
                        }
                    }
                    break;
            }

            // Process sub-groups
            foreach (var subGroup in templateGroup.Groups)
            {
                var tempSubSB = new StringBuilder();
                tempSubSB = ProcessGroup(subGroup, outputSheet);
                tempSB.Append(tempSubSB);
            }

            // Generate Footer
            foreach (var item in templateGroup.Footers)
            {
                // Join items
                tempSB.AppendLine(string.Join(Separator.ToString(), item.Items));
            }

            return tempSB;
        }

        public virtual bool WriteOutput()
        {
            // Check folder to write to exists
            if (WriteSeparateFiles)
            {
                Trace.TraceInformation($"Preparing to write to output folder: {defaultOutputFile}");

                var outDir = new DirectoryInfo(defaultOutputFile);
                if (Directory.Exists(defaultOutputFile))
                {
                    outDir.Clean();
                }
                else
                {
                    outDir.Create();
                }
            }
            else
            {
                Trace.TraceInformation($"Preparing to write output file: {defaultOutputFile}");

                FileInfo fileInfo1 = new FileInfo(defaultOutputFile);

                if (!Directory.Exists(fileInfo1.DirectoryName))
                {
                    Trace.TraceWarning($"Output folder does not exist: {fileInfo1.DirectoryName}. Can't create file.");
                    throw new System.InvalidOperationException($"Output folder does not exist: {fileInfo1.DirectoryName}. Can't create file.");
                }
            }

            if (WriteSeparateFiles)
            {
                // Process each sheet
                foreach (var outputSheet in outputTemplate.OutputSheets)
                {
                    if (outputSheet.IgnoreSheet)
                    {
                        Trace.TraceInformation($"Sheet: {outputSheet.SheetName} is being ignored (via sheet settings).");
                        continue;
                    }

                    string outputFileName = $"{outputSheet.SheetName}{DefaultFileExtension}";
                    string outputFilePath = Path.Combine(defaultOutputFile, outputFileName.Trim());
                    StringBuilder stringBuilder = new StringBuilder();

                    Trace.TraceInformation($"Processing sheet: {outputSheet.SheetName}.");
                    Trace.Indent();

                    // Process the group
                    StringBuilder tempSB = ProcessGroup(outputSheet.MainGroup, outputSheet);

                    // Check if any replacements were not made
                    if (tempSB.ToString().Contains("{"))
                        Trace.TraceWarning($"Remaining '{{' in output sheet: {outputSheet.SheetName}");

                    stringBuilder.Append(tempSB);

                    // Write the separate file
                    Trace.Unindent();
                    Trace.TraceInformation($"Finished sheet: {outputSheet.SheetName}.");
                    Trace.TraceInformation($"Writing to file: {outputFileName}.");

                    StreamWriter outputFile = new StreamWriter(outputFilePath, false, Encoding);
                    outputFile.WriteLine(stringBuilder);
                    outputFile.Close();
                }

                string outputFolder = defaultOutputFile.TrimEnd('\\');
                string sDateTime = $" {DateTime.Now.ToString("yyyy-MM-dd")} {DateTime.Now.ToString("HH-mm-ss")}";
                string outputFolderPlusDateTime = $"{outputFolder}{sDateTime}";

                // Write out extra date/time file if required
                if (WriteTimeStampedCopy)
                {
                    // Copy the whole folder
                    Trace.TraceInformation($"Copying output folder to: {outputFolderPlusDateTime}");

                    try
                    {
                        DirectoryCopier.DirectoryCopy(outputFolder, outputFolderPlusDateTime, true);
                    }
                    catch (Exception)
                    {
                        Trace.TraceWarning($"Unable to copy folder '{outputFolder}' to '{outputFolderPlusDateTime}'.");
                    }
                }
            }

            if (!WriteSeparateFiles)
            {
                // Process each sheet
                StringBuilder stringBuilder = new StringBuilder();
                foreach (var outputSheet in outputTemplate.OutputSheets)
                {
                    if (outputSheet.IgnoreSheet)
                    {
                        Trace.TraceInformation($"Sheet: {outputSheet.SheetName} is being ignored (via sheet settings).");
                        continue;
                    }

                    Trace.TraceInformation($"Processing sheet: {outputSheet.SheetName}.");
                    Trace.Indent();

                    // Process the group
                    StringBuilder tempSB = ProcessGroup(outputSheet.MainGroup, outputSheet);
                    stringBuilder.Append(tempSB);

                    Trace.Unindent();
                    Trace.TraceInformation($"Finished sheet: {outputSheet.SheetName}.");
                }

                // Note that this property does include the file's extension.
                // There does not appear to be a member or property of FileInfo that is just the name without extension.
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(defaultOutputFile);
                FileInfo fileInfo = new FileInfo(defaultOutputFile);

                // Default to writing a single file
                string sDateTime = $" {DateTime.Now.ToString("yyyy-MM-dd")} {DateTime.Now.ToString("HH-mm-ss")}";
                string outputFileName = $"{fileNameWithoutExtension}{fileInfo.Extension}";
                string outputFilePath = Path.Combine(fileInfo.DirectoryName, outputFileName.Trim());
                string outputFileNamePlusDateTime = $"{fileNameWithoutExtension}{sDateTime}{fileInfo.Extension}";
                string outputFilePathPlusDateTime = Path.Combine(fileInfo.DirectoryName, outputFileNamePlusDateTime.Trim());

                // Check if any replacements were not made
                if (stringBuilder.ToString().Contains("{"))
                    Trace.TraceWarning($"Remaining '{{' in output file: {outputFileName}");

                // Write the whole file
                Trace.TraceInformation($"Writing to file: {outputFileName}.");

                StreamWriter outputFile = new StreamWriter(outputFilePath, false, Encoding);
                outputFile.WriteLine(stringBuilder);
                outputFile.Close();

                // Write out extra date/time file if required
                if (WriteTimeStampedCopy)
                {
                    // Copy the whole file
                    Trace.TraceInformation($"Copying output file to: {outputFilePathPlusDateTime}");
                    try
                    {
                        File.Copy(outputFilePath, outputFilePathPlusDateTime);
                    }
                    catch (Exception)
                    {
                        Trace.TraceWarning($"Unable to copy file '{outputFilePath}' to '{outputFilePathPlusDateTime}'.");
                    }
                }
            }

            Trace.TraceInformation($"Finished writing output.");

            return true;
        }

        #endregion Methods
    }

    public class VPLinkTagGenerator : OutputBase2
    {
        #region Constructors

        public VPLinkTagGenerator(SorterToolImporter SorterTool, HelperSettings ProgramSettings)
        {
            sorterTool = SorterTool;
            programSettings = ProgramSettings;

            // CONFIGURATION
            defaultOutputFile = programSettings.VPLinkTagGeneratorOutputFilePath;
            TemplateFilename = "VPLink Tag Generator Template.xlsm";
            WriteTimeStampedCopy = true;
            WriteSeparateFiles = true;
            DefaultFileExtension = ".csv";
            Separator = ',';
            Encoding = Encoding.ASCII;
            // END CONFIGURATION

            // Build up a list of types that can be parsed
            outputTypes = new List<IOutputType>
            {
                new I_O_LIST(this.sorterTool, programSettings)
            };

            if (!string.IsNullOrWhiteSpace(programSettings.OutputTemplateFolder))
            {
                // Read the template
                outputTemplate.ReadXLSXTemplate(Path.Combine(programSettings.OutputTemplateFolder, TemplateFilename));
            }
        }

        #endregion Constructors
    }
}

namespace VPLinkTagGenerator
{
    public class I_O_LIST : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        public Dictionary<string, string> WordReplacments = new Dictionary<string, string>();
        private string _type = "I_O_LIST";

        #endregion Fields

        #region Constructors

        public I_O_LIST(SorterToolImporter sorterTool, HelperSettings programSettings)
        {
            this.sorterTool = sorterTool;
            this.programSettings = programSettings;
        }

        #endregion Constructors

        #region Properties

        public string Type
        {
            get
            {
                return _type;
            }
        }

        #endregion Properties

        #region Methods

        private string PerformReplacements(Dictionary<string, string> WordReplacments, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Replace keywords
            foreach (var item in WordReplacments)
                tempSB = tempSB.Replace(item.Key, item.Value);

            return tempSB;
        }

        private void SetupReplacements(I_O_LISTSheet.I_O_LISTRow ioPoint)
        {
            // Setup replacements
            WordReplacments.Clear();
            WordReplacments.Add("{Tag}", ioPoint.TagnameSafe(true));
            WordReplacments.Add("{DRTagNoSafe}", ioPoint.DRTagNoSafe);
            WordReplacments.Add("{ClientTagNoSafe}", ioPoint.ClientTagNoSafe);
            WordReplacments.Add("{ModulePointSafe}", ioPoint.ModulePointSafe);
            WordReplacments.Add("{Desc}", ioPoint.Description(true));
            WordReplacments.Add("{EngUnits}", ioPoint.EngUnits);
            WordReplacments.Add("{EngHigh}", $"{ioPoint.EngHigh:g}");
            WordReplacments.Add("{EngLow}", $"{ioPoint.EngLow:g}");
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.I_O_LISTSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.SignalType;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (line.Type.Equals(Type) && (string.IsNullOrWhiteSpace(line.SubType) || line.SubType == subtype))
                    {
                        SetupReplacements(item);

                        bool processLine = true;
                        if (!string.IsNullOrWhiteSpace(line.Rule))
                        {
                            // Perform replacements on rule
                            string replacedRule = PerformReplacements(WordReplacments, line.Rule);
                            processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                        }

                        if (processLine)
                        {
                            // Perform replacements
                            string replacedString = PerformReplacements(WordReplacments, string.Join(Separator, line.Items));

                            newSB.AppendLine(replacedString);
                        }
                    }
                }
                count++;
            }

            return newSB;
        }

        public StringBuilder WriteOutputData(string outputSheetName, string RawOutputLine, OutputLine outputLine, bool WriteOneInstanceOnly = false)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();
            StringBuilder tempSB = new StringBuilder(RawOutputLine);

            int count = 1;
            foreach (var item in sorterTool.I_O_LISTSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.SignalType;

                if (string.IsNullOrWhiteSpace(outputLine.SubType) || outputLine.SubType == subtype)
                {
                    SetupReplacements(item);

                    bool processLine = true;
                    if (!string.IsNullOrWhiteSpace(outputLine.Rule))
                    {
                        // Perform replacements on rule
                        string replacedRule = PerformReplacements(WordReplacments, outputLine.Rule);
                        processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                    }

                    if (processLine)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(WordReplacments, RawOutputLine);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }
}