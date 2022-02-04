using Microsoft.VisualStudio.TestTools.UnitTesting;
using SorterToolLibrary.SorterTool;
using System;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace InputTests
{
    [TestClass()]
    public class SorterToolImportExcelTests
    {
        private const string ProductName = "SorterTool";

        #region Fields

        private static Excel.Application app;
        private static Excel.Workbooks workbooks;
        private static Excel.Workbook workbook;

        #endregion Fields

        [TestMethod()]
        public void ReadSorterToolExcelTest()
        {
            SorterToolImporter sorterToolImportExcel = new SorterToolImporter(workbook);
            bool result = sorterToolImportExcel.ReadSorterToolExcel(workbook);

            Assert.IsTrue(result, "SorterTool was not read correctly.");
        }

        [TestMethod()]
        public void ReadSorterToolTest()
        {
            // TODO - Don't remember why we need a DataReader!
            string workbookPath = @"C:\Programming\dotNet\SorterTool\Source\SorterTool\bin\Debug\Master_SorterTool.xlsm";

            SorterToolImporter mySorterToolImport = new SorterToolImporter();
            bool result = mySorterToolImport.ReadSorterToolDataReader(workbookPath);

            Assert.IsTrue(result, "SorterTool was not read correctly.");
        }

        [ClassInitialize]
        public static void TestFixtureSetup(TestContext context)
        {
            // Executes once for the test class. (Optional)

            // Start a log
            string AppDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), ProductName);
            if (!Directory.Exists(AppDataFolder)) Directory.CreateDirectory(AppDataFolder);
            var logfile = Path.Combine(AppDataFolder, $"Application LogFile.log");
            try
            {
                if (File.Exists(logfile))
                    File.Delete(logfile);
            }
            catch (Exception)
            {
                Trace.TraceWarning("Could not delete log file");
                throw new InvalidOperationException($"Could not delete log file");
            }
            Trace.Listeners.Add(new TextWriterTraceListener(logfile));
            Trace.AutoFlush = true;

            app = new Excel.Application();
            workbooks = app.Workbooks;

            //string workbookPath = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory
            //    + @"..\..\..\..\SorterTool\bin\Debug\Master_SorterTool.xlsm");

            string workbookPath = @"C:\Programming\dotNet\SorterTool\Source\SorterTool\bin\Debug\Master_SorterTool.xlsm";
            if (File.Exists(workbookPath))
                workbook = workbooks.Open(workbookPath);
            else
                throw new FileNotFoundException("File not found.", workbookPath);
        }

        [ClassCleanup]
        public static void TestFixtureTearDown()
        {
            // Runs once after all tests in this class are executed. (Optional)
            // Not guaranteed that it executes instantly after all tests from the class.

            if (workbook != null) workbook.Close(SaveChanges: false);
            //if (oldWorkbook != null) oldWorkbook.Close(SaveChanges: false);
            if (workbooks != null) workbooks.Close();
            if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            if (workbooks != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            if (app != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
        }
    }
}