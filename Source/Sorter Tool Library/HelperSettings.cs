using PropertyChanged;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace SorterToolLibrary
{
    [AddINotifyPropertyChangedInterface]
    public class HelperSettings
    {

        #region Properties

        public string ABCodeOutputFolder { get; set; }

        //public string ABMemoryMapOutputFilePath { get; set; }
        public string DRHMIConfigFilePath { get; set; }

        public bool ExportAsSample { get; set; }

        //public string FTViewMEAlarmsOutputFilePath { get; set; }
        //public string ModbusListOutputFilePath { get; set; }
        //public string OPCListOutputFilePath { get; set; }
        public string OutputTemplateFolder { get; set; }
        public string SimaticS7OutputFilePath { get; set; }
        public string SimaticS7OutputSymbolFilePath { get; set; }
        //public string SimaticS7TextLibraryOutputFilePath { get; set; }
        public string SorterToolImportFilepath { get; set; }
        public bool SorterToolUseTagname { get; set; }
        public string TriStationTagOutputTagFilePath { get; set; }
        public string VPLinkTagGeneratorOutputFilePath { get; set; }
        public string WinCC74OutputTagFilePath { get; set; }
        public string WonderwareOutputTagFilePath { get; set; }
        //public string WWTagReplacerOutputFilePath { get; set; }

        #endregion Properties
    }
}