using DRHMIConverter;
using SorterToolLibrary;
using SorterToolLibrary.OutputBase;
using PropertyChanged;
using static DRHMIConverter.OutputTemplate;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using SorterToolLibrary.SorterTool;

namespace SimaticS7SymbolsHelper
{
    public class SimaticS7SymbolsHelper : OutputBase
    {
        #region Fields

        public List<KeywordReplacments> KeyWordReplacements = new List<KeywordReplacments>();

        #endregion Fields

        #region Constructors

        public SimaticS7SymbolsHelper(SorterToolImporter SorterTool, HelperSettings ProgramSettings)
        {
            sorterTool = SorterTool;
            programSettings = ProgramSettings;

            // CONFIGURATION
            defaultOutputFile = programSettings.SimaticS7OutputSymbolFilePath;
            TemplateFilename = "Simatic S7 Symbols Template.xlsx";
            WriteTimeStampedCopy = true;
            WriteSeparateFiles = false;
            DefaultFileExtension = ".sdf";
            Separator = ',';
            Encoding = Encoding.ASCII;
            // END CONFIGURATION

            // Build up a list of types that can be parsed
            outputTypes = new List<IOutputType>
            {
                new DIGIN(SorterTool, programSettings),
                new DIGOUT(SorterTool, programSettings),
                new ANLIN(SorterTool, programSettings),
                new ANLOUT(SorterTool, programSettings),
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

namespace SimaticS7SymbolsHelper
{
    public class DIGIN : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        public Dictionary<string, string> WordReplacments = new Dictionary<string, string>();
        private string _type = "DIGIN";

        #endregion Fields

        #region Constructors

        public DIGIN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(DI_INSheet.DI_INRow diIn, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            // TODO - Make all PerformReplacements a loop like this
            WordReplacments.Clear();
            WordReplacments.Add("{Tag}", diIn.ModulePoint);
            WordReplacments.Add("{Addr}", diIn.Specifier);
            WordReplacments.Add("{CTag}", diIn.ClientTag);
            WordReplacments.Add("{Desc}", diIn.Description);

            // Replace keywords
            foreach (var item in WordReplacments)
                tempSB = tempSB.Replace(item.Key, item.Value);

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var diIn in sorterTool.DI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(diIn, string.Join(Separator, line.Items));

                        // Trim for comment length
                        replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
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
            foreach (var diIn in sorterTool.DI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(diIn, RawOutputLine);

                // Trim for comment length
                replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                newSB.AppendLine(replacedString);

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }

    public class DIGOUT : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        public Dictionary<string, string> WordReplacments = new Dictionary<string, string>();
        private string _type = "DIGOUT";

        #endregion Fields

        #region Constructors

        public DIGOUT(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(DO_INSheet.DO_INRow doIn, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();
            WordReplacments.Add("{Tag}", doIn.ModulePoint);
            WordReplacments.Add("{Addr}", doIn.Specifier);
            WordReplacments.Add("{CTag}", doIn.ClientTag);
            WordReplacments.Add("{Desc}", doIn.Description);

            // Replace keywords
            foreach (var item in WordReplacments)
                tempSB = tempSB.Replace(item.Key, item.Value);

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var diIn in sorterTool.DO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(diIn, string.Join(Separator, line.Items));

                        // Trim for comment length
                        replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
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
            foreach (var diIn in sorterTool.DO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(diIn, RawOutputLine);

                // Trim for comment length
                replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                newSB.AppendLine(replacedString);

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }

    public class ANLIN : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        public Dictionary<string, string> WordReplacments = new Dictionary<string, string>();
        private string _type = "ANLIN";

        #endregion Fields

        #region Constructors

        public ANLIN(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(AI_INSheet.AI_INRow aiIn, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();
            WordReplacments.Add("{Tag}", aiIn.ModulePoint);
            WordReplacments.Add("{Addr}", aiIn.Specifier);
            WordReplacments.Add("{CTag}", aiIn.ClientTag);
            WordReplacments.Add("{Desc}", aiIn.Description);

            // Replace keywords
            foreach (var item in WordReplacments)
                tempSB = tempSB.Replace(item.Key, item.Value);

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var aiIn in sorterTool.AI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(aiIn, string.Join(Separator, line.Items));

                        // Trim for comment length
                        replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
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
            foreach (var aiIn in sorterTool.AI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(aiIn, RawOutputLine);

                // Trim for comment length
                replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                newSB.AppendLine(replacedString);

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }

    public class ANLOUT : IOutputType
    {
        #region Fields

        private HelperSettings programSettings;
        private SorterToolImporter sorterTool;
        public Dictionary<string, string> WordReplacments = new Dictionary<string, string>();
        private string _type = "ANLOUT";

        #endregion Fields

        #region Constructors

        public ANLOUT(SorterToolImporter sorterTool, HelperSettings programSettings)
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

        private string PerformReplacements(AO_INSheet.AO_INRow aoIn, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();
            WordReplacments.Add("{Tag}", aoIn.ModulePoint);
            WordReplacments.Add("{Addr}", aoIn.Specifier);
            WordReplacments.Add("{CTag}", aoIn.ClientTag);
            WordReplacments.Add("{Desc}", aoIn.Description);

            // Replace keywords
            foreach (var item in WordReplacments)
                tempSB = tempSB.Replace(item.Key, item.Value);

            return tempSB;
        }

        public StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator)
        {
            if (sorterTool == null) return new StringBuilder();
            if (programSettings == null) return new StringBuilder();

            StringBuilder newSB = new StringBuilder();

            int count = 1;
            foreach (var aoIn in sorterTool.AO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(aoIn, string.Join(Separator, line.Items));

                        // Trim for comment length
                        replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                        newSB.AppendLine(replacedString);
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
            foreach (var aoIn in sorterTool.AO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(aoIn, RawOutputLine);

                // Trim for comment length
                replacedString = SimaticS7Helper.SimaticS7Helpers.TrimSiemensComment(replacedString);

                newSB.AppendLine(replacedString);

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }
}