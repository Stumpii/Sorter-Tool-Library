using Microsoft.VisualStudio.TestTools.UnitTesting;
using SimaticS7Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutputTests
{
    [TestClass()]
    public class ANALOGTests
    {
        [TestMethod()]
        public void TrimSiemensCommentTest_TooLong()
        {
            string tstInput = "   ANLG_002 : REAL ;	//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A VALUE";
            string expected = "   ANLG_002 : REAL ;	//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A V";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }

        [TestMethod()]
        public void TrimSiemensCommentTest_JustRight()
        {
            string tstInput = "   ANLG_002 : REAL ;	//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A V";
            string expected = "   ANLG_002 : REAL ;	//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A V";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }

        [TestMethod()]
        public void TrimSiemensCommentTest_TooShort()
        {
            string tstInput = "   ANLG_002 : REAL ;	//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE";
            string expected = "   ANLG_002 : REAL ;	//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }

        [TestMethod()]
        public void TrimSiemensCommentTest_JustComment()
        {
            string tstInput = "//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A VALUE";
            string expected = "//LT-5801 1st STAGE SUCTION SCRUBBER LEVEL STAGE SUCTION SCRUBBER LEVEL STAGE A V";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }

        [TestMethod()]
        public void TrimSiemensTitleTest_TooLong()
        {
            string tstInput = "TITLE =TZT-3106 EXTERNAL CONDENSER WATER SUPPLY VALVE OPEN / CLOSE (UV00010) ALM978";
            string expected = "TITLE =TZT-3106 EXTERNAL CONDENSER WATER SUPPLY VALVE OPEN / CLOSE (UV0";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }

        [TestMethod()]
        public void TrimSiemensTitleTest_JustRight()
        {
            string tstInput = "TITLE =TZT-3106 EXTERNAL CONDENSER WATER SUPPLY VALVE OPEN / CLOSE (UV0";
            string expected = "TITLE =TZT-3106 EXTERNAL CONDENSER WATER SUPPLY VALVE OPEN / CLOSE (UV0";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }

        [TestMethod()]
        public void TrimSiemensTitleTest_TooShort()
        {
            string tstInput = "TITLE = my title";
            string expected = "TITLE = my title";
            string result = SimaticS7Helpers.TrimSiemensComment(tstInput);

            Assert.IsTrue(result == expected);
        }
    }
}