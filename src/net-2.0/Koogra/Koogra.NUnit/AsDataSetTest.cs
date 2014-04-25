using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using NUnit.Framework;
using Net.SourceForge.Koogra;

namespace Koogra.NUnit {
    [TestFixture]
    class AsDataSetTest {

        [Test]
        public void AsDataSetXls() {
            IWorkbook workbook = WorkbookFactory.GetExcelBIFFReader(@"..\..\Files\AsDataSet\AsDataSet.xls");
            DataSet set = workbook.AsDataSet(true);
            Assert.AreEqual(1, set.Tables.Count);
            Assert.AreEqual("Sheet1", set.Tables[0].TableName);
            Assert.AreEqual(3, set.Tables[0].Rows.Count);
            Assert.AreEqual(6, set.Tables[0].Columns.Count);

            string[] expectedCols = { "A", "B", "C", "E", "F", "G" };
            for (int i = 0; i < set.Tables[0].Columns.Count; i++) {
                Assert.AreEqual(expectedCols[i], set.Tables[0].Columns[i].ColumnName);
            }
        }

        [Test]
        public void AsDataSetXlsx() {
            IWorkbook workbook = WorkbookFactory.GetExcel2007Reader(@"..\..\Files\AsDataSet\AsDataSet.xlsx");
            DataSet set = workbook.AsDataSet(true);
            Assert.AreEqual(1, set.Tables.Count);
            Assert.AreEqual("Sheet1", set.Tables[0].TableName);
            Assert.AreEqual(3, set.Tables[0].Rows.Count);
            Assert.AreEqual(6, set.Tables[0].Columns.Count);

            string[] expectedCols = { "A", "B", "C", "E", "F", "G" };
            for (int i = 0; i < set.Tables[0].Columns.Count; i++) {
                Assert.AreEqual(expectedCols[i], set.Tables[0].Columns[i].ColumnName);
            }
        }
    }
}
