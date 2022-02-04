using SorterToolLibrary;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SorterToolLibrary.SorterTool;
using System.Diagnostics;
using System.IO;
using System.ComponentModel;
using Settings;

namespace OutputTests
{
    [TestClass()]
    public class SimaticS7HelperTests
    {
        private const string ProductName = "D-R HMI Converter";

        [TestMethod()]
        public void WriteOutputTest()
        {
            string workbookPath = @"C:\Programming\dotNet\SorterTool\Source\SorterTool\bin\Debug\Master_SorterTool.xlsm";

            // Start a log
            string AppDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ProductName);
            var logfile = Path.Combine(AppDataFolder, $"Application LogFile.log");
            try
            {
                if (File.Exists(logfile))
                    File.Delete(logfile);
            }
            catch (Exception)
            {
                Trace.TraceWarning("Could not delete log file");
                throw new System.InvalidOperationException($"Could not delete log file");
            }

            // Read in selected settings
            HelperSettings helpersettings = new TestSettings().GetHelperSettings();
            helpersettings.SorterToolImportFilepath = workbookPath; // TODO - move this to the GetHelperSettings function

            Trace.Listeners.Add(new TextWriterTraceListener(logfile));
            Trace.AutoFlush = true;
            Trace.TraceInformation($"Program settings:");
            foreach (PropertyDescriptor descriptor in TypeDescriptor.GetProperties(helpersettings))
            {
                string name = descriptor.Name;
                object value = descriptor.GetValue(helpersettings);
                Trace.TraceInformation($"\t{name}={value}");
            }

            SorterToolImporter mySorterToolImport = new SorterToolImporter();
            bool result = mySorterToolImport.ReadSorterToolDataReader(helpersettings.SorterToolImportFilepath);

            Assert.IsTrue(result, "SorterTool was not read correctly.");

            SimaticS7Helper.SimaticS7Helper helper = new SimaticS7Helper.SimaticS7Helper(mySorterToolImport, helpersettings);
            bool result2 = helper.WriteOutput();

            Assert.IsTrue(result2, "SimaticS7Helper was not written correctly.");
        }
    }
}