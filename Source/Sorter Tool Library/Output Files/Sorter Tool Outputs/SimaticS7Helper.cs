using SorterToolLibrary;
using SorterToolLibrary.Output_Files.Sorter_Tool_Outputs;
using SorterToolLibrary.OutputBase;
using SorterToolLibrary.SorterTool;
using DRHMIConverter;
using PropertyChanged;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using static DRHMIConverter.OutputTemplate;

namespace SimaticS7Helper
{
    public static class SimaticS7Helpers
    {
        /// <summary>
        /// Truncates a comment (i.e. //abc) to 79 chars in length (max for Siemens)
        /// </summary>
        /// <param name="replacedString"></param>
        /// <returns></returns>
        public static string TrimSiemensComment(string replacedString)
        {
            // TITLE = LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A
            // Does this contain a comment?
            if (replacedString.Contains("//"))
            {
                int slashPos = replacedString.IndexOf("//");
                string comment = replacedString.Substring(slashPos);
                if (comment.Length > 81) // 79 Char total + 2 comment // chars
                {
                    string newComment = comment.Substring(0, 81); // 79 Char total + 2 comment // chars
                    string preComment = replacedString.Substring(0, slashPos);
                    replacedString = preComment + newComment;
                }
            }
            else if (replacedString.StartsWith("TITLE"))
            {
                int equalsPos = replacedString.IndexOf("=");
                string title = replacedString.Substring(equalsPos + 1);
                if (title.Length > 64)
                {
                    string newComment = title.Substring(0, 64);
                    string preComment = replacedString.Substring(0, equalsPos + 1);
                    replacedString = preComment + newComment;
                }
            }

            return replacedString;
        }
    }

    public class SimaticS7Helper : OutputBase
    {
        #region Fields

        public List<KeywordReplacments> KeyWordReplacements = new List<KeywordReplacments>();

        #endregion Fields

        #region Constructors

        public SimaticS7Helper(SorterToolImporter SorterTool, HelperSettings ProgramSettings)
        {
            sorterTool = SorterTool;
            programSettings = ProgramSettings;

            // CONFIGURATION
            defaultOutputFile = programSettings.SimaticS7OutputFilePath;
            TemplateFilename = "Simatic S7 Template.xlsm";
            WriteTimeStampedCopy = true;
            WriteSeparateFiles = true;
            DefaultFileExtension = ".awl";
            Separator = '\t';
            Encoding = Encoding.ASCII;
            // END CONFIGURATION

            // Build up a list of types that can be parsed
            outputTypes = new List<IOutputType>
            {
                new ANLG_IN(SorterTool, programSettings),
                new AO_IN(SorterTool, programSettings),
                new ALM_IN(SorterTool, programSettings),
                new ALM_GEN(SorterTool, programSettings),
                new Bently_In(SorterTool, programSettings),
                new BN_MAP(SorterTool, programSettings),
                new CONFC_IN(SorterTool, programSettings),
                new Cust_MB1_MAP(SorterTool, programSettings),
                new Cust_MB2_MAP(SorterTool, programSettings),
                new Cust_MB3_MAP(SorterTool, programSettings),
                new Cust_MB4_MAP(SorterTool, programSettings),
                new Cust_MB5_MAP(SorterTool, programSettings),
                new Cust_MB6_MAP(SorterTool, programSettings),
                new Cust_MB7_MAP(SorterTool, programSettings),
                new Cust_MB8_MAP(SorterTool, programSettings),
                new Cust_SDV_GEN(SorterTool, programSettings),
                new Cust_SDNV_GEN(SorterTool, programSettings),
                new Cust_STD_DEV_SEL_3(SorterTool, programSettings),
                new Cust_V2oo3Bypass(SorterTool, programSettings),
                new MB1_MAP(SorterTool, programSettings),
                new MB2_MAP(SorterTool, programSettings),
                new MB3_MAP(SorterTool, programSettings),
                new MB4_MAP(SorterTool, programSettings),
                new MB5_MAP(SorterTool, programSettings),
                new MB6_MAP(SorterTool, programSettings),
                new MB7_MAP(SorterTool, programSettings),
                new MB8_MAP(SorterTool, programSettings),
                new SD_GEN(SorterTool, programSettings),
                new SD_IN(SorterTool, programSettings),
                new STAT_IN(SorterTool, programSettings),
                new STD_DEV_SEL_3(SorterTool, programSettings),
                new TIMERS_IN(SorterTool, programSettings),
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
                    throw new System.InvalidOperationException($"Template does not exist: {template}");
                }
            }
            else
            {
                Trace.TraceWarning($"OutputTemplateFolder does not exist: {programSettings.OutputTemplateFolder}");
                throw new System.InvalidOperationException($"OutputTemplateFolder does not exist: {programSettings.OutputTemplateFolder}");
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

namespace SimaticS7Helper
{
    public class ALM_GEN : IOutputType
    {
        #region Fields

        private readonly string _type = "ALM_GEN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

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

        #endregion Methods
    }

    /// <summary>
    /// Class representing an output type for WinCC
    /// </summary>
    public class ALM_IN : IOutputType
    {
        #region Fields

        private readonly string _type = "ALM_IN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public ALM_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.ALM_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.ALM_INSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(ALM_INSheet.ALM_INRow alarm, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{alarm.Index:000}");

            // Text
            tempSB = tempSB.Replace("{AlarmNumText}", alarm.AlarmRef);
            tempSB = tempSB.Replace("{Tag}", alarm.AlarmTag);
            tempSB = tempSB.Replace("{Message}", alarm.AlarmDescription);
            tempSB = tempSB.Replace("{Type}", alarm.AlarmType);
            tempSB = tempSB.Replace("{CustomPLCTag}", alarm.CustomPLCTag);

            return tempSB;
        }

        #endregion Methods
    }

    public class ANLG_IN : IOutputType
    {
        #region Fields

        private readonly string _type = "ANLG_IN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public ANLG_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.ANLG_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";
                //if (!string.IsNullOrWhiteSpace(item.Address))
                subtype = "USED_ANLG";
                //subtype = item.Source;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.ANLG_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

                if (!string.IsNullOrWhiteSpace(item.Address))
                    subtype = "USED_ANLG";

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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(ANLG_INSheet.ANLG_INRow analog, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{analog.Index:000}");

            // Text
            tempSB = tempSB.Replace("{Tag}", analog.ClientTag);
            tempSB = tempSB.Replace("{Description}", analog.Description);
            tempSB = tempSB.Replace("{Address}", analog.Address);

            tempSB = tempSB.Replace("{EngMin}", $"{analog.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{analog.EngMax:e}");

            tempSB = tempSB.Replace("{SFLoLimit}", $"{analog.SFLoLimit:e}");
            tempSB = tempSB.Replace("{SFHiLimit}", $"{analog.SFHiLimit:e}");
            tempSB = tempSB.Replace("{LoAlmLimit}", $"{analog.LoAlmLimit:e}");
            tempSB = tempSB.Replace("{LoSDLimit}", $"{analog.LoSDLimit:e}");
            tempSB = tempSB.Replace("{HiAlmLimit}", $"{analog.HiAlmLimit:e}");
            tempSB = tempSB.Replace("{HiSDLimit}", $"{analog.HiSDLimit:e}");
            tempSB = tempSB.Replace("{OpenLimit}", $"{analog.OpenLimit:e}");

            return tempSB;
        }

        #endregion Methods
    }

    public class AO_IN : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        private readonly string _type = "AO_IN";

        #endregion Fields

        #region Constructors

        public AO_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(AO_INSheet.AO_INRow aoin, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{aoin.Index:000}");

            // Text
            tempSB = tempSB.Replace("{Tag}", aoin.ClientTag);
            tempSB = tempSB.Replace("{Description}", aoin.Description);
            tempSB = tempSB.Replace("{Address}", aoin.Address);

            tempSB = tempSB.Replace("{EngMin}", $"{aoin.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{aoin.EngMax:e}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.AO_INSheet.Rows)
            {
                if (string.IsNullOrWhiteSpace(item.Address)) continue;

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.AO_INSheet.Rows)
            {
                if (string.IsNullOrWhiteSpace(item.Address)) continue;

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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

    public class Bently_In : IOutputType
    {
        #region Fields

        private readonly string _type = "Bently_In";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Bently_In(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.Bently_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.Bently_INSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(Bently_INSheet.Bently_InRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{item.Index:000}");

            // Text
            tempSB = tempSB.Replace("{Description}", item.Description);
            tempSB = tempSB.Replace("{ValueAddr}", item.ValueAddress);
            tempSB = tempSB.Replace("{StatusAddr}", item.StatusAddress);
            tempSB = tempSB.Replace("{GapAddr}", item.GapAddress);
            tempSB = tempSB.Replace("{HH_SP_Addr}", item.HH_SP_Addr);
            tempSB = tempSB.Replace("{H_SP_Addr}", item.H_SP_Addr);
            tempSB = tempSB.Replace("{L_SP_Addr}", item.L_SP_Addr);
            tempSB = tempSB.Replace("{LL_SP_Addr}", item.LL_SP_Addr);
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class BN_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "BN_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public BN_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.BN_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.BN_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }
                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(BN_MAPSheet.BN_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{SlotRef}", item.ModulePoint);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddress);
            tempSB = tempSB.Replace("{Description}", item.I_O_List?.RawDescription ?? "");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB1_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB1_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB1_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB1_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB1_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB2_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB2_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB2_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB2_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB2_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB3_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB3_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB3_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB3_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB3_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB4_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB4_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB4_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB4_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB4_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB5_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB5_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB5_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB5_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB5_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB6_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB6_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB6_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB6_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB6_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB7_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB7_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB7_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB7_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB7_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class MB8_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "MB8_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public MB8_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB8_MAPSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.FCSubType;

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.MB8_MAPSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{DBName}", item.DBName);
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);
            tempSB = tempSB.Replace("{DBNum}", $"{item.SIEDBNum:0}");
            tempSB = tempSB.Replace("{PLCTagname}", $"{item.PLCTagname:0}");
            tempSB = tempSB.Replace("{IndexNumPlus1}", $"{item.IndexNumPlus1:0}");
            tempSB = tempSB.Replace("{IndexNum}", $"{item.IndexNum:0}");
            tempSB = tempSB.Replace("{EngMin}", $"{item.EngMin:e}");
            tempSB = tempSB.Replace("{EngMax}", $"{item.EngMax:e}");
            tempSB = tempSB.Replace("{AnlgNum}", $"{item.AnlgNum:000}");

            return tempSB;
        }

        #endregion Methods
    }

    public class Cust_MB1_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB1_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB1_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB1_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB1_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB1_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB2_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB2_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB2_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB2_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB2_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB2_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB3_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB3_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB3_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB3_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB3_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB1_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB4_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB4_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB4_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB4_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB4_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB4_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB5_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB5_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB5_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB5_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB5_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB5_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB6_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB6_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB6_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB6_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB6_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB6_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB7_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB7_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB7_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB7_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB7_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB7_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class Cust_MB8_MAP : IOutputType
    {
        #region Fields

        private readonly string _type = "Cust_MB8_MAP";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public Cust_MB8_MAP(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.MB8_MAPSheet.Rows)
            {
                //Trace.TraceInformation($"Processing Index{item.Index:000}");

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = item.DBSubType;

                //Trace.TraceInformation($"Processing Type: {item.GenType}");

                if (item.DBSubType == "INT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        INTUnpacked iNTUnpacked = new INTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB8_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "INT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        iNTUnpacked.PLCAddr = row.PLCAddr;
                                        iNTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        iNTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        iNTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        iNTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        iNTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        iNTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        iNTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        iNTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        iNTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        iNTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        iNTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        iNTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        iNTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        iNTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        iNTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        iNTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        iNTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsINT_B(iNTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsINT_B(iNTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else if (item.DBSubType == "DINT_B")
                {
                    // Only pick up the first row that has PLC address
                    if ((item.PLCAddr ?? "") != "")
                    {
                        //Trace.TraceInformation($"Processing Type: {item.PLCAddr}");

                        // Class to store DINT data
                        DINTUnpacked dINTUnpacked = new DINTUnpacked();

                        bool inWord = false;
                        string currAddr = "";
                        foreach (var row in sorterTool.MB8_MAPSheet.Rows)
                        {
                            inWord = (row.DBSubType == "DINT_B" && row.PLCAddr == item.PLCAddr) || inWord;
                            if (inWord)
                            {
                                if (currAddr == "")
                                    currAddr = row.PLCAddr;

                                // If this is the start of a new PLC address, then the current address is done.
                                if (row.PLCAddr != "" && row.PLCAddr != currAddr)
                                    break;

                                switch (row.BitAddr)
                                {
                                    case "Bit00":
                                        dINTUnpacked.PLCAddr = row.PLCAddr;
                                        dINTUnpacked.ClientTag_Bit00 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit00 = row.ClientDescription;
                                        break;

                                    case "Bit01":
                                        dINTUnpacked.ClientTag_Bit01 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit01 = row.ClientDescription;
                                        break;

                                    case "Bit02":
                                        dINTUnpacked.ClientTag_Bit02 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit02 = row.ClientDescription;
                                        break;

                                    case "Bit03":
                                        dINTUnpacked.ClientTag_Bit03 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit03 = row.ClientDescription;
                                        break;

                                    case "Bit04":
                                        dINTUnpacked.ClientTag_Bit04 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit04 = row.ClientDescription;
                                        break;

                                    case "Bit05":
                                        dINTUnpacked.ClientTag_Bit05 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit05 = row.ClientDescription;
                                        break;

                                    case "Bit06":
                                        dINTUnpacked.ClientTag_Bit06 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit06 = row.ClientDescription;
                                        break;

                                    case "Bit07":
                                        dINTUnpacked.ClientTag_Bit07 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit07 = row.ClientDescription;
                                        break;

                                    case "Bit08":
                                        dINTUnpacked.ClientTag_Bit08 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit08 = row.ClientDescription;
                                        break;

                                    case "Bit09":
                                        dINTUnpacked.ClientTag_Bit09 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit09 = row.ClientDescription;
                                        break;

                                    case "Bit10":
                                        dINTUnpacked.ClientTag_Bit10 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit10 = row.ClientDescription;
                                        break;

                                    case "Bit11":
                                        dINTUnpacked.ClientTag_Bit11 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit11 = row.ClientDescription;
                                        break;

                                    case "Bit12":
                                        dINTUnpacked.ClientTag_Bit12 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit12 = row.ClientDescription;
                                        break;

                                    case "Bit13":
                                        dINTUnpacked.ClientTag_Bit13 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit13 = row.ClientDescription;
                                        break;

                                    case "Bit14":
                                        dINTUnpacked.ClientTag_Bit14 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit14 = row.ClientDescription;
                                        break;

                                    case "Bit15":
                                        dINTUnpacked.ClientTag_Bit15 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit15 = row.ClientDescription;
                                        break;

                                    case "Bit16":
                                        dINTUnpacked.ClientTag_Bit16 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit16 = row.ClientDescription;
                                        break;

                                    case "Bit17":
                                        dINTUnpacked.ClientTag_Bit17 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit17 = row.ClientDescription;
                                        break;

                                    case "Bit18":
                                        dINTUnpacked.ClientTag_Bit18 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit18 = row.ClientDescription;
                                        break;

                                    case "Bit19":
                                        dINTUnpacked.ClientTag_Bit19 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit19 = row.ClientDescription;
                                        break;

                                    case "Bit20":
                                        dINTUnpacked.ClientTag_Bit20 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit20 = row.ClientDescription;
                                        break;

                                    case "Bit21":
                                        dINTUnpacked.ClientTag_Bit21 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit21 = row.ClientDescription;
                                        break;

                                    case "Bit22":
                                        dINTUnpacked.ClientTag_Bit22 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit22 = row.ClientDescription;
                                        break;

                                    case "Bit23":
                                        dINTUnpacked.ClientTag_Bit23 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit23 = row.ClientDescription;
                                        break;

                                    case "Bit24":
                                        dINTUnpacked.ClientTag_Bit24 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit24 = row.ClientDescription;
                                        break;

                                    case "Bit25":
                                        dINTUnpacked.ClientTag_Bit25 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit25 = row.ClientDescription;
                                        break;

                                    case "Bit26":
                                        dINTUnpacked.ClientTag_Bit26 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit26 = row.ClientDescription;
                                        break;

                                    case "Bit27":
                                        dINTUnpacked.ClientTag_Bit27 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit27 = row.ClientDescription;
                                        break;

                                    case "Bit28":
                                        dINTUnpacked.ClientTag_Bit28 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit28 = row.ClientDescription;
                                        break;

                                    case "Bit29":
                                        dINTUnpacked.ClientTag_Bit29 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit29 = row.ClientDescription;
                                        break;

                                    case "Bit30":
                                        dINTUnpacked.ClientTag_Bit30 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit30 = row.ClientDescription;
                                        break;

                                    case "Bit31":
                                        dINTUnpacked.ClientTag_Bit31 = row.ClientTag;
                                        dINTUnpacked.ClientDescription_Bit31 = row.ClientDescription;
                                        break;

                                    default:
                                        break;
                                }
                            }
                        }

                        Trace.TraceInformation($"Writing output");

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
                                    string replacedRule = PerformReplacementsDINT_B(dINTUnpacked, line.Rule);
                                    processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                                }

                                if (processLine)
                                {
                                    // Perform replacements
                                    string replacedString = PerformReplacementsDINT_B(dINTUnpacked, string.Join(Separator, line.Items));

                                    // Trim for comment length
                                    replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                    newSB.AppendLine(replacedString);
                                }
                            }
                        }
                    }
                }
                else
                {
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

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        private string PerformReplacements(MB_MAPRow item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{BitAddr}", item.BitAddr);
            tempSB = tempSB.Replace("{ClientTag}", item.ClientTag);
            tempSB = tempSB.Replace("{ClientDescription}", item.ClientDescription);

            return tempSB;
        }

        private string PerformReplacementsDINT_B(DINTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientTag_Bit16}", item.ClientTag_Bit16);
            tempSB = tempSB.Replace("{ClientTag_Bit17}", item.ClientTag_Bit17);
            tempSB = tempSB.Replace("{ClientTag_Bit18}", item.ClientTag_Bit18);
            tempSB = tempSB.Replace("{ClientTag_Bit19}", item.ClientTag_Bit19);
            tempSB = tempSB.Replace("{ClientTag_Bit20}", item.ClientTag_Bit20);
            tempSB = tempSB.Replace("{ClientTag_Bit21}", item.ClientTag_Bit21);
            tempSB = tempSB.Replace("{ClientTag_Bit22}", item.ClientTag_Bit22);
            tempSB = tempSB.Replace("{ClientTag_Bit23}", item.ClientTag_Bit23);
            tempSB = tempSB.Replace("{ClientTag_Bit24}", item.ClientTag_Bit24);
            tempSB = tempSB.Replace("{ClientTag_Bit25}", item.ClientTag_Bit25);
            tempSB = tempSB.Replace("{ClientTag_Bit26}", item.ClientTag_Bit26);
            tempSB = tempSB.Replace("{ClientTag_Bit27}", item.ClientTag_Bit27);
            tempSB = tempSB.Replace("{ClientTag_Bit28}", item.ClientTag_Bit28);
            tempSB = tempSB.Replace("{ClientTag_Bit29}", item.ClientTag_Bit29);
            tempSB = tempSB.Replace("{ClientTag_Bit30}", item.ClientTag_Bit30);
            tempSB = tempSB.Replace("{ClientTag_Bit31}", item.ClientTag_Bit31);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit16}", item.ClientDescription_Bit16);
            tempSB = tempSB.Replace("{ClientDescription_Bit17}", item.ClientDescription_Bit17);
            tempSB = tempSB.Replace("{ClientDescription_Bit18}", item.ClientDescription_Bit18);
            tempSB = tempSB.Replace("{ClientDescription_Bit19}", item.ClientDescription_Bit19);
            tempSB = tempSB.Replace("{ClientDescription_Bit20}", item.ClientDescription_Bit20);
            tempSB = tempSB.Replace("{ClientDescription_Bit21}", item.ClientDescription_Bit21);
            tempSB = tempSB.Replace("{ClientDescription_Bit22}", item.ClientDescription_Bit22);
            tempSB = tempSB.Replace("{ClientDescription_Bit23}", item.ClientDescription_Bit23);
            tempSB = tempSB.Replace("{ClientDescription_Bit24}", item.ClientDescription_Bit24);
            tempSB = tempSB.Replace("{ClientDescription_Bit25}", item.ClientDescription_Bit25);
            tempSB = tempSB.Replace("{ClientDescription_Bit26}", item.ClientDescription_Bit26);
            tempSB = tempSB.Replace("{ClientDescription_Bit27}", item.ClientDescription_Bit27);
            tempSB = tempSB.Replace("{ClientDescription_Bit28}", item.ClientDescription_Bit28);
            tempSB = tempSB.Replace("{ClientDescription_Bit29}", item.ClientDescription_Bit29);
            tempSB = tempSB.Replace("{ClientDescription_Bit30}", item.ClientDescription_Bit30);
            tempSB = tempSB.Replace("{ClientDescription_Bit31}", item.ClientDescription_Bit31);

            return tempSB;
        }

        private string PerformReplacementsINT_B(INTUnpacked item, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Text
            tempSB = tempSB.Replace("{PLCAddr}", item.PLCAddr);
            tempSB = tempSB.Replace("{ClientTag_Bit00}", item.ClientTag_Bit00);
            tempSB = tempSB.Replace("{ClientTag_Bit01}", item.ClientTag_Bit01);
            tempSB = tempSB.Replace("{ClientTag_Bit02}", item.ClientTag_Bit02);
            tempSB = tempSB.Replace("{ClientTag_Bit03}", item.ClientTag_Bit03);
            tempSB = tempSB.Replace("{ClientTag_Bit04}", item.ClientTag_Bit04);
            tempSB = tempSB.Replace("{ClientTag_Bit05}", item.ClientTag_Bit05);
            tempSB = tempSB.Replace("{ClientTag_Bit06}", item.ClientTag_Bit06);
            tempSB = tempSB.Replace("{ClientTag_Bit07}", item.ClientTag_Bit07);
            tempSB = tempSB.Replace("{ClientTag_Bit08}", item.ClientTag_Bit08);
            tempSB = tempSB.Replace("{ClientTag_Bit09}", item.ClientTag_Bit09);
            tempSB = tempSB.Replace("{ClientTag_Bit10}", item.ClientTag_Bit10);
            tempSB = tempSB.Replace("{ClientTag_Bit11}", item.ClientTag_Bit11);
            tempSB = tempSB.Replace("{ClientTag_Bit12}", item.ClientTag_Bit12);
            tempSB = tempSB.Replace("{ClientTag_Bit13}", item.ClientTag_Bit13);
            tempSB = tempSB.Replace("{ClientTag_Bit14}", item.ClientTag_Bit14);
            tempSB = tempSB.Replace("{ClientTag_Bit15}", item.ClientTag_Bit15);
            tempSB = tempSB.Replace("{ClientDescription_Bit00}", item.ClientDescription_Bit00);
            tempSB = tempSB.Replace("{ClientDescription_Bit01}", item.ClientDescription_Bit01);
            tempSB = tempSB.Replace("{ClientDescription_Bit02}", item.ClientDescription_Bit02);
            tempSB = tempSB.Replace("{ClientDescription_Bit03}", item.ClientDescription_Bit03);
            tempSB = tempSB.Replace("{ClientDescription_Bit04}", item.ClientDescription_Bit04);
            tempSB = tempSB.Replace("{ClientDescription_Bit05}", item.ClientDescription_Bit05);
            tempSB = tempSB.Replace("{ClientDescription_Bit06}", item.ClientDescription_Bit06);
            tempSB = tempSB.Replace("{ClientDescription_Bit07}", item.ClientDescription_Bit07);
            tempSB = tempSB.Replace("{ClientDescription_Bit08}", item.ClientDescription_Bit08);
            tempSB = tempSB.Replace("{ClientDescription_Bit09}", item.ClientDescription_Bit09);
            tempSB = tempSB.Replace("{ClientDescription_Bit10}", item.ClientDescription_Bit10);
            tempSB = tempSB.Replace("{ClientDescription_Bit11}", item.ClientDescription_Bit11);
            tempSB = tempSB.Replace("{ClientDescription_Bit12}", item.ClientDescription_Bit12);
            tempSB = tempSB.Replace("{ClientDescription_Bit13}", item.ClientDescription_Bit13);
            tempSB = tempSB.Replace("{ClientDescription_Bit14}", item.ClientDescription_Bit14);
            tempSB = tempSB.Replace("{ClientDescription_Bit15}", item.ClientDescription_Bit15);

            return tempSB;
        }

        #endregion Methods

        internal class DINTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientTag_Bit16 { get; set; } = "";
            public string ClientTag_Bit17 { get; set; } = "";
            public string ClientTag_Bit18 { get; set; } = "";
            public string ClientTag_Bit19 { get; set; } = "";
            public string ClientTag_Bit20 { get; set; } = "";
            public string ClientTag_Bit21 { get; set; } = "";
            public string ClientTag_Bit22 { get; set; } = "";
            public string ClientTag_Bit23 { get; set; } = "";
            public string ClientTag_Bit24 { get; set; } = "";
            public string ClientTag_Bit25 { get; set; } = "";
            public string ClientTag_Bit26 { get; set; } = "";
            public string ClientTag_Bit27 { get; set; } = "";
            public string ClientTag_Bit28 { get; set; } = "";
            public string ClientTag_Bit29 { get; set; } = "";
            public string ClientTag_Bit30 { get; set; } = "";
            public string ClientTag_Bit31 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string ClientDescription_Bit16 { get; set; } = "";
            public string ClientDescription_Bit17 { get; set; } = "";
            public string ClientDescription_Bit18 { get; set; } = "";
            public string ClientDescription_Bit19 { get; set; } = "";
            public string ClientDescription_Bit20 { get; set; } = "";
            public string ClientDescription_Bit21 { get; set; } = "";
            public string ClientDescription_Bit22 { get; set; } = "";
            public string ClientDescription_Bit23 { get; set; } = "";
            public string ClientDescription_Bit24 { get; set; } = "";
            public string ClientDescription_Bit25 { get; set; } = "";
            public string ClientDescription_Bit26 { get; set; } = "";
            public string ClientDescription_Bit27 { get; set; } = "";
            public string ClientDescription_Bit28 { get; set; } = "";
            public string ClientDescription_Bit29 { get; set; } = "";
            public string ClientDescription_Bit30 { get; set; } = "";
            public string ClientDescription_Bit31 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }

        internal class INTUnpacked
        {
            public string ClientTag_Bit00 { get; set; } = "";
            public string ClientTag_Bit01 { get; set; } = "";
            public string ClientTag_Bit02 { get; set; } = "";
            public string ClientTag_Bit03 { get; set; } = "";
            public string ClientTag_Bit04 { get; set; } = "";
            public string ClientTag_Bit05 { get; set; } = "";
            public string ClientTag_Bit06 { get; set; } = "";
            public string ClientTag_Bit07 { get; set; } = "";
            public string ClientTag_Bit08 { get; set; } = "";
            public string ClientTag_Bit09 { get; set; } = "";
            public string ClientTag_Bit10 { get; set; } = "";
            public string ClientTag_Bit11 { get; set; } = "";
            public string ClientTag_Bit12 { get; set; } = "";
            public string ClientTag_Bit13 { get; set; } = "";
            public string ClientTag_Bit14 { get; set; } = "";
            public string ClientTag_Bit15 { get; set; } = "";
            public string ClientDescription_Bit00 { get; set; } = "";
            public string ClientDescription_Bit01 { get; set; } = "";
            public string ClientDescription_Bit02 { get; set; } = "";
            public string ClientDescription_Bit03 { get; set; } = "";
            public string ClientDescription_Bit04 { get; set; } = "";
            public string ClientDescription_Bit05 { get; set; } = "";
            public string ClientDescription_Bit06 { get; set; } = "";
            public string ClientDescription_Bit07 { get; set; } = "";
            public string ClientDescription_Bit08 { get; set; } = "";
            public string ClientDescription_Bit09 { get; set; } = "";
            public string ClientDescription_Bit10 { get; set; } = "";
            public string ClientDescription_Bit11 { get; set; } = "";
            public string ClientDescription_Bit12 { get; set; } = "";
            public string ClientDescription_Bit13 { get; set; } = "";
            public string ClientDescription_Bit14 { get; set; } = "";
            public string ClientDescription_Bit15 { get; set; } = "";
            public string PLCAddr { get; set; } = "";
        }
    }

    public class CONFC_IN : IOutputType
    {
        #region Fields

        private readonly string _type = "CONFC_IN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public CONFC_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.CONFC_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.CONFC_INSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(CONFC_INSheet.CONFC_INRow confc, string RawOutputLine)
        {
            string outputLine = RawOutputLine;

            // Perform replacements
            outputLine = outputLine.Replace("{Index}", $"{confc.Index:000}");

            // Text
            outputLine = outputLine.Replace("{Tag}", confc.Tag.Replace(" ", "_"));
            outputLine = outputLine.Replace("{Description}", confc.Description);

            return outputLine;
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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

    public class SD_IN : IOutputType
    {
        #region Fields

        private readonly string _type = "SD_IN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public SD_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.SD_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.SD_INSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(SD_INSheet.SD_INRow alarm, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{alarm.AlarmNumber:000}");

            // Text
            tempSB = tempSB.Replace("{AlarmNumText}", alarm.AlarmRef);
            tempSB = tempSB.Replace("{Tag}", alarm.AlarmTag);
            tempSB = tempSB.Replace("{Message}", alarm.AlarmDescription);
            tempSB = tempSB.Replace("{Type}", alarm.AlarmType);
            tempSB = tempSB.Replace("{CustomPLCTag}", alarm.CustomPLCTag);

            return tempSB;
        }

        #endregion Methods
    }

    public class STAT_IN : IOutputType
    {
        #region Fields

        private readonly string _type = "STAT_IN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public STAT_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.STAT_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.STAT_INSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(STAT_INSheet.STAT_INRow stat, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{stat.Index:000}");

            // Text
            tempSB = tempSB.Replace("{Tag}", stat.PLCTag);
            tempSB = tempSB.Replace("{Description}", stat.Description);

            return tempSB;
        }

        #endregion Methods
    }

    public class TIMERS_IN : IOutputType
    {
        #region Fields

        private readonly string _type = "TIMERS_IN";
        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;

        #endregion Fields

        #region Constructors

        public TIMERS_IN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.TIMERS_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.TIMERS_INSheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
                    }
                }

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        private string PerformReplacements(TIMERS_INSheet.TIMERS_INRow timer, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{timer.Index:000}");

            // Text
            tempSB = tempSB.Replace("{Description}", timer.Description);
            tempSB = tempSB.Replace("{Preset}", (timer.Preset ?? 0).ToString());

            return tempSB;
        }

        #endregion Methods
    }

    public class Cust_V2oo3Bypass : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        private readonly string _type = "Cust_V2oo3Bypass";

        #endregion Fields

        #region Constructors

        public Cust_V2oo3Bypass(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(List<ALM_GENSheet.ALM_GENRow> aLM_GENs, List<ALM_GENSheet.ALM_GENRow> aLM_GEN_Flts, SD_GENSheet.SD_GENRow sdGen, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            #region Lookup Alarm Info

            if (aLM_GENs.Count != 3)
            {
                Trace.TraceWarning($"When generating {Type}, did not find 3 matching alarms in the ALM_GEN sheet.");
                return tempSB;
            }
            if (aLM_GEN_Flts.Count != 3)
            {
                Trace.TraceWarning($"When generating {Type}, did not find 3 matching alarms in the ALM_GEN sheet.");
                return tempSB;
            }

            #endregion Lookup Alarm Info

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{sdGen.Index:000}");
            tempSB = tempSB.Replace("{AlarmNumber}", $"{sdGen.AlarmNumber:000}");
            tempSB = tempSB.Replace("{A_Trip}", $"{aLM_GENs[0].Index:000}");
            tempSB = tempSB.Replace("{B_Trip}", $"{aLM_GENs[1].Index:000}");
            tempSB = tempSB.Replace("{C_Trip}", $"{aLM_GENs[2].Index:000}");
            tempSB = tempSB.Replace("{A_Fault}", $"{aLM_GEN_Flts[0].Index:000}");
            tempSB = tempSB.Replace("{B_Fault}", $"{aLM_GEN_Flts[1].Index:000}");
            tempSB = tempSB.Replace("{C_Fault}", $"{aLM_GEN_Flts[2].Index:000}");
            tempSB = tempSB.Replace("{A_Bypass}", $"{aLM_GENs[0].Index:000}");
            tempSB = tempSB.Replace("{B_Bypass}", $"{aLM_GENs[1].Index:000}");
            tempSB = tempSB.Replace("{C_Bypass}", $"{aLM_GENs[2].Index:000}");
            tempSB = tempSB.Replace("{Trip}", $"{sdGen.AlarmNumber:000}");
            tempSB = tempSB.Replace("{Description}", $"{sdGen.SD_IN?.AlarmDescription ?? ""}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (SD_GENSheet.SD_GENRow sdGenData in sorterTool.SD_GENSheet.Rows)
            {
                if (!string.IsNullOrWhiteSpace(sdGenData.VotingGroup))
                {
                    // Find matching alarms with matching voting group and type (HH or LL)
                    List<ALM_GENSheet.ALM_GENRow> aLM_GEN_Alms = new List<ALM_GENSheet.ALM_GENRow>();
                    foreach (ALM_GENSheet.ALM_GENRow almGenData in sorterTool.ALM_GENSheet.Rows)
                    {
                        if (almGenData.VotingGroup == sdGenData.VotingGroup &&
                            almGenData.AlarmType == sdGenData.AlarmType.Replace("_Voted", ""))
                        {
                            aLM_GEN_Alms.Add(almGenData);
                        }
                    }

                    // Find matching signal fail for each analog
                    List<ALM_GENSheet.ALM_GENRow> aLM_GEN_Flts = new List<ALM_GENSheet.ALM_GENRow>();
                    foreach (ALM_GENSheet.ALM_GENRow sfItem in aLM_GEN_Alms)
                    {
                        foreach (ALM_GENSheet.ALM_GENRow almGenData in sorterTool.ALM_GENSheet.Rows)
                        {
                            if (almGenData.AnlgStatIndex == sfItem.AnlgStatIndex &&
                                almGenData.AlarmType == "SF")
                            {
                                aLM_GEN_Flts.Add(almGenData);
                            }
                        }
                    }

                    if (programSettings.ExportAsSample && count > 5) break;

                    // Define subtype for this item, or blank to generate anyway
                    string subtype = "";

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
                                string replacedRule = PerformReplacements(aLM_GEN_Alms, aLM_GEN_Flts, sdGenData, line.SubType);
                                processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                            }

                            if (processLine)
                            {
                                // Perform replacements
                                string replacedString = PerformReplacements(aLM_GEN_Alms, aLM_GEN_Flts, sdGenData, string.Join(Separator, line.Items));

                                // Trim for comment length
                                replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

                                newSB.AppendLine(replacedString);
                            }
                        }
                    }
                    count++;
                }
            }

            return newSB;
        }

        public StringBuilder WriteOutputData(string outputSheetName, string RawOutputLine, OutputLine outputLine, bool WriteOneInstanceOnly = false)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        #endregion Methods
    }

    public class Cust_STD_DEV_SEL_3 : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        private readonly string _type = "Cust_STD_DEV_SEL_3";

        #endregion Fields

        #region Constructors

        public Cust_STD_DEV_SEL_3(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(STD_DEV_SEL_3Sheet.STD_DEV_SEL_3Row sTD_DEV_SEL_3, List<ALM_GENSheet.ALM_GENRow> aLM_GENs, List<ALM_GENSheet.ALM_GENRow> aLM_GEN_Flts, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{sTD_DEV_SEL_3.Index:000}");
            tempSB = tempSB.Replace("{A_Input}", $"{aLM_GENs[0].AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{B_Input}", $"{aLM_GENs[1].AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{C_Input}", $"{aLM_GENs[2].AnlgStatIndex:000}");
            tempSB = tempSB.Replace("{A_Fault}", $"{aLM_GEN_Flts[0].Index:000}");
            tempSB = tempSB.Replace("{B_Fault}", $"{aLM_GEN_Flts[1].Index:000}");
            tempSB = tempSB.Replace("{C_Fault}", $"{aLM_GEN_Flts[2].Index:000}");
            tempSB = tempSB.Replace("{Sel_Anlg}", $"{sTD_DEV_SEL_3.SEL_ANLG:000}");
            tempSB = tempSB.Replace("{SF_DIR}", $"{sTD_DEV_SEL_3.SF_DIR}");
            tempSB = tempSB.Replace("{SEL_MODE}", $"{sTD_DEV_SEL_3.SEL_MODE}");
            tempSB = tempSB.Replace("{A_Description}", $"{aLM_GENs[0].ALM_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{B_Description}", $"{aLM_GENs[1].ALM_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{C_Description}", $"{aLM_GENs[2].ALM_IN?.AlarmDescription ?? ""}");
            tempSB = tempSB.Replace("{VotingGroup}", $"{sTD_DEV_SEL_3.VotingGroup}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var stdDevSel3InData in sorterTool.STD_DEV_SEL_3Sheet.Rows)
            {
                if (stdDevSel3InData.VotingGroup == "")
                    continue;

                // Check if any analogs have voting
                int votedAnalogs = 0;
                foreach (var item in sorterTool.CAESheet.Rows)
                {
                    if (item.VotingGroup == stdDevSel3InData.VotingGroup)
                        votedAnalogs += 1;
                }
                if (votedAnalogs == 0)
                {
                    Trace.TraceWarning($"For voting group {stdDevSel3InData.VotingGroup} in STD_DEV_SEL_3Sheet, " +
                        $"there are no rows in CAE that have that voting group.");
                    continue;
                }

                // Find matching alarms with matching voting group and type (HH or LL).
                List<ALM_GENSheet.ALM_GENRow> aLM_GEN_Alms = new List<ALM_GENSheet.ALM_GENRow>();
                foreach (ALM_GENSheet.ALM_GENRow almGenData in sorterTool.ALM_GENSheet.Rows)
                {
                    if (almGenData.VotingGroup == stdDevSel3InData.VotingGroup &&
                        almGenData.AlarmType == "DEV")
                    {
                        aLM_GEN_Alms.Add(almGenData);
                    }
                }

                // Find matching signal fail for each analog
                List<ALM_GENSheet.ALM_GENRow> aLM_GEN_Flts = new List<ALM_GENSheet.ALM_GENRow>();
                foreach (ALM_GENSheet.ALM_GENRow sfItem in aLM_GEN_Alms)
                {
                    foreach (ALM_GENSheet.ALM_GENRow almGenData in sorterTool.ALM_GENSheet.Rows)
                    {
                        if (almGenData.AnlgStatIndex == sfItem.AnlgStatIndex &&
                            almGenData.AlarmType == "SF")
                        {
                            aLM_GEN_Flts.Add(almGenData);
                        }
                    }
                }

                #region Lookup Alarm Info

                if (aLM_GEN_Alms.Count != 3)
                {
                    Trace.TraceWarning($"When generating {Type}, did not find 3 matching alarms in the ALM_GEN sheet.");
                    continue;
                }
                if (aLM_GEN_Flts.Count != 3)
                {
                    Trace.TraceWarning($"When generating {Type}, did not find 3 matching alarms in the ALM_GEN sheet.");
                    continue;
                }

                #endregion Lookup Alarm Info

                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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
                            string replacedRule = PerformReplacements(stdDevSel3InData, aLM_GEN_Alms, aLM_GEN_Flts, line.SubType);
                            processLine = SimpleExpressionEvaluator.EvaluateUsingCalcEngine(replacedRule);
                        }

                        if (processLine)
                        {
                            // Perform replacements
                            string replacedString = PerformReplacements(stdDevSel3InData, aLM_GEN_Alms, aLM_GEN_Flts, string.Join(Separator, line.Items));

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

            Trace.TraceWarning($"Typ {Type} does not support GROUPBY_OUTPUT or GROUPBY_SINGLE order. Must be GROUPBY_INPUT order.");

            return newSB;
        }

        #endregion Methods
    }

    public class SD_GEN : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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

    public class STD_DEV_SEL_3 : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        private readonly string _type = "STD_DEV_SEL_3";

        #endregion Fields

        #region Constructors

        public STD_DEV_SEL_3(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(STD_DEV_SEL_3Sheet.STD_DEV_SEL_3Row sTD_DEV_SEL_3, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Perform replacements
            tempSB = tempSB.Replace("{Index}", $"{sTD_DEV_SEL_3.Index:000}");
            tempSB = tempSB.Replace("{Sel_Anlg}", $"{sTD_DEV_SEL_3.SEL_ANLG:000}");
            tempSB = tempSB.Replace("{SF_DIR}", $"{sTD_DEV_SEL_3.SF_DIR}");
            tempSB = tempSB.Replace("{SEL_MODE}", $"{sTD_DEV_SEL_3.SEL_MODE}");

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var item in sorterTool.STD_DEV_SEL_3Sheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Define subtype for this item, or blank to generate anyway
                string subtype = "";

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

                            // Trim for comment length
                            replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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
            foreach (var item in sorterTool.STD_DEV_SEL_3Sheet.Rows)
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

                        // Trim for comment length
                        replacedString = SimaticS7Helpers.TrimSiemensComment(replacedString);

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