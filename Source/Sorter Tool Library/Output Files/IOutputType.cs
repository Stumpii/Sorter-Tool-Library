using DRHMIConverter;
using static DRHMIConverter.OutputTemplate;
using System.Text;
using SorterToolLibrary.Output_Files.Sorter_Tool_Outputs;

namespace DRHMIConverter
{
    /// <summary>
    /// Defines an interface for an output template type (i.e. ALM_IN, AI_IN, TIMER)
    /// </summary>
    public interface IOutputType
    {
        #region "Properties"

        string Type { get; }

        #endregion "Properties"

        #region "Methods"

        StringBuilder WriteInputData(TemplateGroup outputGroup, string Separator);

        StringBuilder WriteOutputData(string outputSheetName, string RawOutputLine, OutputLine outputLine, bool WriteOneInstanceOnly = false);

        #endregion "Methods"
    }
}