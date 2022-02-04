using SorterToolLibrary;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SorterToolLibrary.SorterTool;
using Settings;

namespace OutputTests
{
    [TestClass()]
    public class SimaticS7SymbolsTests
    {
        [TestMethod()]
        public void WriteOutputTest()
        {
            string workbookPath = @"C:\Programming\dotNet\SorterTool\Source\SorterTool\bin\Debug\Master_SorterTool.xlsm";

            // Read in selected settings
            HelperSettings helpersettings = new TestSettings().GetHelperSettings();
            helpersettings.SorterToolImportFilepath = workbookPath; // TODO - move this to the GetHelperSettings function

            SorterToolImporter mySorterToolImport = new SorterToolImporter();
            bool result = mySorterToolImport.ReadSorterToolDataReader(helpersettings.SorterToolImportFilepath);

            Assert.IsTrue(result, "SorterTool was not read correctly.");

            SimaticS7SymbolsHelper.SimaticS7SymbolsHelper symbolhelper = new SimaticS7SymbolsHelper.SimaticS7SymbolsHelper(mySorterToolImport, helpersettings);
            bool result2 = symbolhelper.WriteOutput();

            Assert.IsTrue(result2, "SimaticS7SymbolsHelper was not written correctly.");
        }
    }
}