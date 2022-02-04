using Microsoft.VisualStudio.TestTools.UnitTesting;
using SorterToolLibrary.SorterTool;

namespace InputTests
{
    [TestClass()]
    public class I_O_LISTRowTests
    {
        [TestMethod()]
        public void SplitStringTest()
        {
            string test;
            string[] expected;
            string[] result;

            // tt[t]-[r]ss/cc (DI-01/03), tt[t]-[r]ss_cc (DI-01_03) or tt[t][r]ss_CHcc (DI01_CH03)

            test = "DI-01/03";
            expected = new string[] { "DI", "-", "01", "/", "03" };
            result = ExcelHelpers.SplitString(test);
            Assert.AreEqual(string.Join(",", result), string.Join(",", expected));

            test = "DI-_01-/03";
            expected = new string[] { "DI", "-_", "01", "-/", "03" };
            result = ExcelHelpers.SplitString(test);
            Assert.AreEqual(string.Join(",", result), string.Join(",", expected));

            test = "DI-01_03";
            expected = new string[] { "DI", "-", "01", "_", "03" };
            result = ExcelHelpers.SplitString(test);
            Assert.AreEqual(string.Join(",", result), string.Join(",", expected));

            test = "DI01_CH03";
            expected = new string[] { "DI", "01", "_", "CH", "03" };
            result = ExcelHelpers.SplitString(test);
            Assert.AreEqual(string.Join(",", result), string.Join(",", expected));
        }
    }
}