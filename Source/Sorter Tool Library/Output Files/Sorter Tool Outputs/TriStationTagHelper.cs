using DRHMIConverter;
using SorterToolLibrary;
using SorterToolLibrary.OutputBase;
using SorterToolLibrary.SorterTool;
using static DRHMIConverter.OutputTemplate;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace TriStationTagHelper
{
    public class TriStationTagHelper : OutputBase
    {
        #region Constructors

        public TriStationTagHelper(SorterToolImporter SorterTool, HelperSettings ProgramSettings)
        {
            sorterTool = SorterTool;
            programSettings = ProgramSettings;

            // CONFIGURATION
            defaultOutputFile = programSettings.TriStationTagOutputTagFilePath;
            TemplateFilename = "TriStation Tag Template.xlsx";
            WriteTimeStampedCopy = true;
            WriteSeparateFiles = false;
            DefaultFileExtension = "";
            Separator = ',';
            Encoding = Encoding.ASCII;
            // END CONFIGURATION

            // Build up a list of types that can be parsed
            outputTypes = new List<IOutputType>
            {
                new DIGIN(this.sorterTool, programSettings),
                new DIGOUT(this.sorterTool, programSettings),
                new ANLIN(this.sorterTool, programSettings),
                new ANLOUT(this.sorterTool, programSettings)
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

namespace TriStationTagHelper
{
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

        private string PerformReplacements(AI_INSheet.AI_INRow ioPoint, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();

            if (ioPoint.I_O_List == null)
                return tempSB;

            WordReplacments.Add("{Tag}", ioPoint.I_O_List.TagnameSafe(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Desc}", ioPoint.I_O_List.Description(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Group1}", $"{ioPoint.I_O_List.SignalType}_{ioPoint.I_O_List.RackSlot}_{ioPoint.I_O_List.Channel}");
            WordReplacments.Add("{Alias}", Convert.ToString(ioPoint.I_O_List.BitANLG + 30000));
            WordReplacments.Add("{Slot}", ioPoint.I_O_List.Slot);
            WordReplacments.Add("{Point}", ioPoint.I_O_List.Channel);

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
            foreach (var item in sorterTool.AI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));
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
            foreach (var item in sorterTool.AI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(item, RawOutputLine);
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

        private string PerformReplacements(AO_INSheet.AO_INRow ioPoint, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();

            if (ioPoint.I_O_List == null)
                return tempSB;

            WordReplacments.Add("{Tag}", ioPoint.I_O_List.TagnameSafe(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Desc}", ioPoint.I_O_List.Description(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Group1}", $"{ioPoint.SignalType}_{ioPoint.I_O_List.RackSlot}_{ioPoint.I_O_List.Channel}");
            WordReplacments.Add("{Alias}", Convert.ToString(ioPoint.I_O_List.BitANLG + 40000));
            WordReplacments.Add("{Slot}", ioPoint.I_O_List.Slot);
            WordReplacments.Add("{Point}", ioPoint.I_O_List.Channel);

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
            foreach (var item in sorterTool.AO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));
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

            foreach (var item in sorterTool.AO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(item, RawOutputLine);
                newSB.AppendLine(replacedString);

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }

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

        private string PerformReplacements(DI_INSheet.DI_INRow ioPoint, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();

            if (ioPoint.I_O_List == null)
                return tempSB;

            WordReplacments.Add("{Tag}", ioPoint.I_O_List.TagnameSafe(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Desc}", ioPoint.I_O_List.Description(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Group1}", $"{ioPoint.SignalType}_{ioPoint.I_O_List.RackSlot}_{ioPoint.I_O_List.Channel}");
            WordReplacments.Add("{Alias}", Convert.ToString(ioPoint.I_O_List.BitANLG + 10000));
            WordReplacments.Add("{Slot}", ioPoint.I_O_List.Slot);
            WordReplacments.Add("{Point}", ioPoint.I_O_List.Channel);

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
            foreach (var item in sorterTool.DI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));
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
            foreach (var item in sorterTool.DI_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(item, RawOutputLine);
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

        private string PerformReplacements(DO_INSheet.DO_INRow ioPoint, string RawOutputLine)
        {
            string tempSB = RawOutputLine;

            // Setup replacements
            WordReplacments.Clear();

            if (ioPoint.I_O_List == null)
                return tempSB;

            WordReplacments.Add("{Tag}", ioPoint.I_O_List.TagnameSafe(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Desc}", ioPoint.I_O_List.Description(programSettings.SorterToolUseTagname));
            WordReplacments.Add("{Group1}", $"{ioPoint.SignalType}_{ioPoint.I_O_List.RackSlot}_{ioPoint.I_O_List.Channel}");
            WordReplacments.Add("{Alias}", Convert.ToString(ioPoint.I_O_List.BitANLG + 0));
            WordReplacments.Add("{Slot}", ioPoint.I_O_List.Slot);
            WordReplacments.Add("{Point}", ioPoint.I_O_List.Channel);

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
            foreach (var item in sorterTool.DO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Process all lines
                foreach (var line in outputGroup.Data)
                {
                    // Check if this is the right output type (i.e. ANLG)
                    if (Type == line.Type)
                    {
                        // Perform replacements
                        string replacedString = PerformReplacements(item, string.Join(Separator, line.Items));
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
            foreach (var item in sorterTool.DO_INSheet.Rows)
            {
                if (programSettings.ExportAsSample && count > 5) break;

                // Perform replacements
                string replacedString = PerformReplacements(item, RawOutputLine);
                newSB.AppendLine(replacedString);

                count++;
                if (WriteOneInstanceOnly) break;
            }
            return newSB;
        }

        #endregion Methods
    }
}