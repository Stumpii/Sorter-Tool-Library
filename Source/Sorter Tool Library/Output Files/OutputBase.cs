using DRHMIConverter;
using static DRHMIConverter.OutputTemplate;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SorterToolLibrary.SorterTool;

namespace SorterToolLibrary.OutputBase
{
    /// <summary>
    /// Base class for an output template (i.e. SimaticS7Helper) that defines the options and types handled by the template.
    /// </summary>
    public abstract class OutputBase
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
        /// Enable the output file/folder name to be a file path
        /// </summary>
        public bool AllowOutputFile { get; set; } = true;

        /// <summary>
        /// Allow the output file/folder name to be a folder path
        /// </summary>
        public bool AllowOutputFolder { get; set; } = false;

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

        internal bool CheckOutputFileFolder()
        {
            bool folderexist = Directory.Exists(defaultOutputFile);
            bool fileExists = File.Exists(defaultOutputFile);

            FileInfo fileInfo1 = new FileInfo(defaultOutputFile);
            bool filesFolderExists = Directory.Exists(fileInfo1.DirectoryName);

            if (!AllowOutputFile)
            {
                if (fileExists)
                {
                    Trace.TraceWarning($"Output file selected. Select output folder instead");
                    return false;
                }
            }
            if (!AllowOutputFolder)
            {
                if (folderexist)
                {
                    Trace.TraceWarning($"Output folder selected. Select output file instead");
                    return false;
                }
            }

            return true;
        }

        public virtual bool WriteOutput()
        {
            // Check output file/folder selection
            bool checkOkay = CheckOutputFileFolder();
            if (!checkOkay) return false;

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
                    stringBuilder.Append(tempSB);

                    // Check if any replacements were not made
                    if (tempSB.ToString().Contains("{"))
                        Trace.TraceWarning($"Remaining '{{' in output sheet: {outputSheet.SheetName}");

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
}