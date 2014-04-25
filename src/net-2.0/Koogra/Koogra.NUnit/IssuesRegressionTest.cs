using System;
using NUnit.Framework;
using Net.SourceForge.Koogra.Excel;

namespace Net.SourceForge.Koogra.NUnit
{
    /// <summary>
    /// This class if used for regression testing Koogra bugs.
    /// </summary>
    [TestFixture()]
    public class IssueRegressionTest
    {
        /// <summary>
        /// NUnit requires a default constructor.
        /// </summary>
        public IssueRegressionTest()
        {
        }

        [Test()]
        public void FarEastTest()
        {
            DumpWorkbookToConsole("far east.xls");
        }

        /// <summary>
        /// This issue was reported by Chris Bell. This test needs the index problem.xls file.
        /// </summary>
        [Test()]
        public void IndexProblemTest()
        {
            DumpWorkbookToConsole(@"..\..\Files\IndexProblem\Index Problem.xls");
        }

        [Test()]
        public void JapanSampleTest()
        {
            DumpWorkbookToConsole(@"..\..\Files\JapanSample\book4.xls");
        }

        /// <summary>
        /// This issue was reported by  Daniel Rusk. This test needs the rk problem.xls file.
        /// </summary>
        [Test()]
        public void RKProblemIssue()
        {
            Workbook wb = new Workbook(@"..\..\Files\RKProblem\rk problem.xls");
            Worksheet ws = wb.Sheets[0];

            decimal value = Convert.ToDecimal(ws.Rows[1].Cells[4].Value);
            Assert.AreEqual(914551.7d, value);
        }

        /// <summary>
        /// This issue was reported by Chris Bell. This test needs the APTOptionPricer USA Example.xls file.
        /// </summary>
        [Test()]
        public void IndexOutOfBoundsIssue()
        {
            DumpWorkbookToConsole(@"..\..\Files\IndexOutOfBounds\APTOptionPricer USA Example.xls");
        }

        [Test()]
        public void BoolErrSwitchErrorTest()
        {
            DumpWorkbookToConsole(@"..\..\Files\BoolErrorSwitchError\BoolErrorSwitchError.xls");
        }

        [Test()]
        public void WorksheetNameEncodingISsue()
        {
            Workbook wb = new Workbook("17876.xls");


            foreach (Worksheet ws in wb.Sheets)
                Console.WriteLine(ws.Name);
        }

        /// <summary>
        /// This method just exercises the excel workbook data.
        /// </summary>
        /// <param name="path">The path to the workbook.</param>
        private Workbook DumpWorkbookToConsole(string path)
        {
            // print the path
            Console.WriteLine(path);

            // construct our workbook
            Workbook wb = new Workbook(path);

            // dump the worksheet data
            foreach (Worksheet ws in wb.Sheets)
            {
                Console.Write("Sheet is ");
                Console.Write(ws.Name);

                Console.Write(" First row is: ");
                Console.Write(ws.Rows.MinRow);

                Console.Write(" Last row is: ");
                Console.WriteLine(ws.Rows.MaxRow);

                // dump cell data
                for (uint r = ws.Rows.MinRow; r <= ws.Rows.MaxRow; ++r)
                {
                    Row row = ws.Rows[r];

                    if (row != null)
                    {
                        Console.Write("Row: ");
                        Console.Write(r);

                        Console.Write(" First Col: ");
                        Console.Write(row.Cells.MinCol);

                        Console.Write(" Last Col: ");
                        Console.WriteLine(row.Cells.MaxCol);

                        for (uint c = row.Cells.MinCol; c <= row.Cells.MaxCol; ++c)
                        {
                            Cell cell = row.Cells[c];

                            Console.Write("Col: ");
                            Console.Write(c);

                            if (cell != null)
                            {
                                Console.Write(" Value: ");
                                Console.Write(cell.Value);
                                Console.Write(" Formatted Value: ");
                                Console.WriteLine(cell.FormattedValue());
                            }
                            else
                                Console.WriteLine(" null");
                        }
                    }
                }
            }

            return wb;
        }

        /// <summary>
        /// This issue was reported by Chris Bell. This test needs the APTxVAR COM Server Component Monte Carlo VaR Example.xls file.
        /// </summary>
        [Test()]
        public void InvalidTypeCastIssue()
        {
            DumpWorkbookToConsole(@"..\..\Files\InvalidTypeCast\APTxVAR COM Server Component Monte Carlo VaR Example.xls");
        }

        /// <summary>
        /// This issue was reported by Wilson Chan. Koogra was never tested for reading Excel files from the Far East.
        /// Bug was in the SstRecord fill method. The far east data size was being read as 2 bytes even though it should have been read as 4 bytes.
        /// </summary>
        [Test()]
        public void SimpleFarEastCharacterIssue()
        {
            Workbook wb = DumpWorkbookToConsole(@"..\..\Files\SimpleFarEastCharacters\Book1.xls");

            Worksheet ws = wb.Sheets[0];

            Row row = ws.Rows[1];

            Assert.IsTrue(row.Cells[0].Value.ToString().IndexOf('\x0') < 0);

            row = ws.Rows[2];

            Assert.IsNotNull(row.Cells[0].Value);
        }
    }
}
