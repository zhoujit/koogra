using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Net.SourceForge.Koogra.Excel2007;

namespace Koogra.NUnit
{
    [TestFixture()]
  public  class Issues2007RegressionTest
    {
        public Issues2007RegressionTest()
        {

        }

        [Test()]
        public void Issue()
        {
            var wb = new Workbook("VT Test File 1.xlsx");

            var ws = wb.GetWorksheetByName("Settings");

            var r = ws.GetRow(12);

            var c = r.GetCell(0);

            Console.WriteLine(c.Value);
            Assert.AreEqual("Cue card message given", c.Value);
        }
    }
}
