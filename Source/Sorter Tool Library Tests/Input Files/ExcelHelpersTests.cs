using Microsoft.VisualStudio.TestTools.UnitTesting;
using SorterToolLibrary.SorterTool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InputTests
{
    [TestClass()]
    public class ExcelHelpersTests
    {
        [TestMethod()]
        public void ColumnLetterTest()
        {
            string res = ExcelHelpers.ColumnLetter(0, true);
            Assert.IsTrue(res == "A");

            res = ExcelHelpers.ColumnLetter(77, true);
            Assert.IsTrue(res == "BZ");

            res = ExcelHelpers.ColumnLetter(78, true);
            Assert.IsTrue(res == "");

            res = ExcelHelpers.ColumnLetter(1, false);
            Assert.IsTrue(res == "A");

            res = ExcelHelpers.ColumnLetter(78, false);
            Assert.IsTrue(res == "BZ");

            res = ExcelHelpers.ColumnLetter(79, false);
            Assert.IsTrue(res == "");
        }
    }
}