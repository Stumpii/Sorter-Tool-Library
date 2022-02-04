using DRHMIConverter;
using SorterToolLibrary;
using SorterToolLibrary.OutputBase;
using SorterToolLibrary.Output_Files.Sorter_Tool_Outputs;
using SorterToolLibrary.SorterTool;
using PropertyChanged;
using static DRHMIConverter.OutputTemplate;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace ABGenerator
{
    public class ABGenerator : OutputBase
    {
        #region Fields

        public List<KeywordReplacments> KeyWordReplacements = new List<KeywordReplacments>();

        #endregion Fields

        #region Constructors

        public ABGenerator(SorterToolImporter SorterTool, HelperSettings ProgramSettings)
        {
            sorterTool = SorterTool;
            programSettings = ProgramSettings;

            // CONFIGURATION
            defaultOutputFile = programSettings.ABCodeOutputFolder;
            TemplateFilename = "AB Code Template.xlsm";
            WriteTimeStampedCopy = true;
            WriteSeparateFiles = true;
            DefaultFileExtension = ".txt";
            Separator = '\t';
            Encoding = Encoding.ASCII;
            // END CONFIGURATION

            // Build up a list of types that can be parsed
            outputTypes = new List<IOutputType>
            {
                new ALM_GEN(SorterTool, programSettings),
                new Cust_SDNV_GEN(SorterTool, programSettings),
                new Cust_SDV_GEN(SorterTool, programSettings),
                new SD_GEN(SorterTool, programSettings),
            };

            // TODO - make these changes to all other conversions
            if (Directory.Exists(programSettings.OutputTemplateFolder))
            {
                // Read the template
                string template = Path.Combine(programSettings.OutputTemplateFolder, TemplateFilename);

                if (File.Exists(template))
                {
                    Trace.TraceInformation($"Reading template: {template}");
                    outputTemplate.ReadXLSXTemplate(template);
                }
                else
                {
                    Trace.TraceWarning($"Template does not exist: {template}");
                    throw new InvalidOperationException($"Template does not exist: {template}");
                }
            }
            else
            {
                Trace.TraceWarning($"OutputTemplateFolder does not exist: {programSettings.OutputTemplateFolder}");
                throw new InvalidOperationException($"OutputTemplateFolder does not exist: {programSettings.OutputTemplateFolder}");
            }
        }

        #endregion Constructors

        #region Classes

        [AddINotifyPropertyChangedInterface]
        public class KeywordReplacments
        {
            #region Fields

            public Dictionary<string, string> WordReplacments = new Dictionary<string, string>();

            #endregion Fields

            #region Properties

            public string Type { get; set; }

            #endregion Properties
        }

        #endregion Classes
    }
}

namespace ABGenerator
{
    public class ALM_GEN : IOutputType
    {
        #region Fields

        private readonly HelperSettings programSettings;
        private readonly SorterToolImporter sorterTool;
        private readonly string _type = "ALM_GEN";

        #endregion Fields

        #region Constructors

        public ALM_GEN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(ALM_GENSheet.ALM_GENRow almgen, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            #region Lookup STAT Info

            string Tag = string.Empty;

            // Only check for ANLG/STAT point if one specified
            if ((almgen.AnlgStatIndex ?? 0) > 0)
            {
                STAT_INSheet.STAT_INRow statIn = sorterTool.FindSTAT_INRow(almgen.AnlgStatIndex);
                if (statIn != null)
                    Tag = statIn.PLCTag;
                else
                    Trace.TraceWarning($"During replacement in {Type}, could not find instance in STAT_IN for {almgen}.");
            }

            #endregion Lookup STAT Info

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{almgen.Index:000}");
            tempSB = tempSB.Replace("{AnlgStatIndex}", $"{almgen.AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{Tag}", Tag);
            tempSB = tempSB.Replace("{Message}", $"{almgen.ALM_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{BentlyStatus}", $"{almgen.BentlyStatus:000}");
            tempSB = tempSB.Replace("{TimerIndex}", $"{almgen.Index + 500}");
            tempSB = tempSB.Replace("{VotingGroup}", $"{almgen.VotingGroup}");
            tempSB = tempSB.Replace("{CustomPLCTag}", $"{almgen.ALM_IN?.CustomPLCTag ?? ""}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.ALM_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = ($"{item.IOType}_{item.AlarmType}_{item.AlarmClass}").Replace("__", "_").Trim('_');

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (line.Type.Equals(Type) && (string.IsNullOrWhiteSpace(line.SubType) || line.SubType == subtype))
                    {
                        bool processLine = true;
                        if (!string.IsNullOrWhiteSpace(line.Rule))
                        {
                            // Perform replacements on rule
                            string replacedRule = PerformReplacements(item, line.Rule);
                            processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                        }

                        if (processLine)
                        {
                            // Perform replacements
                            string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));

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

            int count = 1;
            foreach (var item in sorterTool.ALM_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

                if (string.IsNullOrWhiteSpace(outputLine.SubType) || outputLine.SubType == subtype)
                {
                    bool processLine = true;
                    if (!string.IsNullOrWhiteSpace(outputLine.Rule))
                    {
                        // Perform replacements on rule
                        string replacedRule = PerformReplacements(item, outputLine.Rule);
                        processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                    }

                    if (processLine)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, RawOutputLine);

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

    public class Cust_SDNV_GEN : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        private readonly string _type = "Cust_SDNV_GEN";

        #endregion Fields

        #region Constructors

        public Cust_SDNV_GEN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(SD_GENSheet.SD_GENRow sdgen, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            #region Lookup STAT Info

            string Tag = string.Empty;

            // Only check for ANLG/STAT point if one specified
            // TODO Move this out to the calling routine and pass in one object holding all the info.
            if ((sdgen.AnlgStatIndex ?? 0) > 0)
            {
                STAT_INSheet.STAT_INRow statIn = sorterTool.FindSTAT_INRow(sdgen.AnlgStatIndex);
                if (statIn != null)
                    Tag = statIn.PLCTag;
                else
                    Trace.TraceWarning($"During replacement in {Type}, could not find instance in STAT_IN for {sdgen}.");
            }

            #endregion Lookup STAT Info

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{sdgen.Index:000}");
            tempSB = tempSB.Replace("{AlarmNumber}", $"{sdgen.AlarmNumber:000}");
            tempSB = tempSB.Replace("{AnlgStatIndex}", $"{sdgen.AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{Tag}", Tag);
            tempSB = tempSB.Replace("{Message}", $"{sdgen.SD_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{BentlyStatus}", $"{sdgen.BentlyStatus:000}");
            tempSB = tempSB.Replace("{TimerIndex}", $"{sdgen.AlarmNumber + 500}");
            tempSB = tempSB.Replace("{CustomPLCTag}", $"{sdgen.SD_IN?.CustomPLCTag ?? ""}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.SD_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Skip if the shutdown is vented
                if (item.IsVented) continue;

                // Define subtype for this item, or blank to generate anyway
                string subtype = ($"{item.IOType}_{item.AlarmType}_{item.AlarmClass}").Replace("__", "_").Trim('_');

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (line.Type.Equals(Type) && (string.IsNullOrWhiteSpace(line.SubType) || line.SubType == subtype))
                    {
                        bool processLine = true;
                        if (!string.IsNullOrWhiteSpace(line.Rule))
                        {
                            // Perform replacements on rule
                            string replacedRule = PerformReplacements(item, line.Rule);
                            processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                        }

                        if (processLine)
                        {
                            // Perform replacements
                            string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));

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

            int count = 1;
            foreach (var item in sorterTool.SD_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

                // Skip if the shutdown is vented
                if (item.IsVented) continue;

                if (string.IsNullOrWhiteSpace(outputLine.SubType) || outputLine.SubType == subtype)
                {
                    bool processLine = true;
                    if (!string.IsNullOrWhiteSpace(outputLine.Rule))
                    {
                        // Perform replacements on rule
                        string replacedRule = PerformReplacements(item, outputLine.Rule);
                        processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                    }

                    if (processLine)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, RawOutputLine);

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

    public class Cust_SDV_GEN : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        private readonly string _type = "Cust_SDV_GEN";

        #endregion Fields

        #region Constructors

        public Cust_SDV_GEN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(SD_GENSheet.SD_GENRow sdgen, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            #region Lookup STAT Info

            string Tag = string.Empty;

            // Only check for ANLG/STAT point if one specified
            if ((sdgen.AnlgStatIndex ?? 0) > 0)
            {
                STAT_INSheet.STAT_INRow statIn = sorterTool.FindSTAT_INRow(sdgen.AnlgStatIndex);
                if (statIn != null)
                    Tag = statIn.PLCTag;
                else
                    Trace.TraceWarning($"During replacement in {Type}, could not find instance in STAT_IN for {sdgen}.");
            }

            #endregion Lookup STAT Info

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{sdgen.Index:000}");
            tempSB = tempSB.Replace("{AlarmNumber}", $"{sdgen.AlarmNumber:000}");
            tempSB = tempSB.Replace("{AnlgStatIndex}", $"{sdgen.AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{Tag}", Tag);
            tempSB = tempSB.Replace("{Message}", $"{sdgen.SD_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{BentlyStatus}", $"{sdgen.BentlyStatus:000}");
            tempSB = tempSB.Replace("{TimerIndex}", $"{sdgen.AlarmNumber + 500}");
            tempSB = tempSB.Replace("{CustomPLCTag}", $"{sdgen.SD_IN?.CustomPLCTag ?? ""}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.SD_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Skip if the shutdown is non-vented
                if (!item.IsVented) continue;

                // Define subtype for this item, or blank to generate anyway
                string subtype = ($"{item.IOType}_{item.AlarmType}_{item.AlarmClass}").Replace("__", "_").Trim('_');

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (line.Type.Equals(Type) && (string.IsNullOrWhiteSpace(line.SubType) || line.SubType == subtype))
                    {
                        bool processLine = true;
                        if (!string.IsNullOrWhiteSpace(line.Rule))
                        {
                            // Perform replacements on rule
                            string replacedRule = PerformReplacements(item, line.Rule);
                            processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                        }

                        if (processLine)
                        {
                            // Perform replacements
                            string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));

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

            int count = 1;
            foreach (var item in sorterTool.SD_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

                // Skip if the shutdown is non-vented
                if (!item.IsVented) continue;

                if (string.IsNullOrWhiteSpace(outputLine.SubType) || outputLine.SubType == subtype)
                {
                    bool processLine = true;
                    if (!string.IsNullOrWhiteSpace(outputLine.Rule))
                    {
                        // Perform replacements on rule
                        string replacedRule = PerformReplacements(item, outputLine.Rule);
                        processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                    }

                    if (processLine)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, RawOutputLine);

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

    public class SD_GEN : IOutputType
    {
        #region Fields

        private readonly HelperSettings programSettings;
        private readonly SorterToolImporter sorterTool;
        private readonly string _type = "SD_GEN";

        #endregion Fields

        #region Constructors

        public SD_GEN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(SD_GENSheet.SD_GENRow almgen, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            #region Lookup STAT Info

            string Tag = string.Empty;

            // Only check for ANLG/STAT point if one specified
            if ((almgen.AnlgStatIndex ?? 0) > 0)
            {
                STAT_INSheet.STAT_INRow statIn = sorterTool.FindSTAT_INRow(almgen.AnlgStatIndex);
                if (statIn != null)
                    Tag = statIn.PLCTag;
                else
                    Trace.TraceWarning($"During replacement in {Type}, could not find instance in STAT_IN for {almgen}.");
            }

            #endregion Lookup STAT Info

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{almgen.Index:000}");
            tempSB = tempSB.Replace("{AlarmNumber}", $"{almgen.AlarmNumber:000}");
            tempSB = tempSB.Replace("{AnlgStatIndex}", $"{almgen.AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{Tag}", Tag);
            tempSB = tempSB.Replace("{Message}", $"{almgen.SD_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{BentlyStatus}", $"{almgen.BentlyStatus:000}");
            tempSB = tempSB.Replace("{TimerIndex}", $"{almgen.AlarmNumber + 500}");
            tempSB = tempSB.Replace("{VotingGroup}", $"{almgen.VotingGroup}");
            tempSB = tempSB.Replace("{CustomPLCTag}", $"{almgen.SD_IN?.CustomPLCTag ?? ""}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.SD_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = ($"{item.IOType}_{item.AlarmType}_{item.AlarmClass}").Replace("__", "_").Trim('_');

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (line.Type.Equals(Type) && (string.IsNullOrWhiteSpace(line.SubType) || line.SubType == subtype))
                    {
                        bool processLine = true;
                        if (!string.IsNullOrWhiteSpace(line.Rule))
                        {
                            // Perform replacements on rule
                            string replacedRule = PerformReplacements(item, line.Rule);
                            processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                        }

                        if (processLine)
                        {
                            // Perform replacements
                            string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));

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

            int count = 1;
            foreach (var item in sorterTool.SD_GENSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

                if (string.IsNullOrWhiteSpace(outputLine.SubType) || outputLine.SubType == subtype)
                {
                    bool processLine = true;
                    if (!string.IsNullOrWhiteSpace(outputLine.Rule))
                    {
                        // Perform replacements on rule
                        string replacedRule = PerformReplacements(item, outputLine.Rule);
                        processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                    }

                    if (processLine)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, RawOutputLine);

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