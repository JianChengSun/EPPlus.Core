using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlus.Core.FunctionalTests.SampleApp
{
    [TestClass]
    public class Sample17
    {
        [TestMethod]
        public void RunSample17()
        {
            using (var fs = new FileStream(Path.Combine("bin", "check.xlsx"), FileMode.Open, FileAccess.Read))
            using (var package = new ExcelPackage(fs))
            {
                foreach (var worksheet1 in package.Workbook.Worksheets)
                {
                    var prCollection = worksheet1.ProtectedRanges;
                    if (prCollection.Count != 1)
                        throw new InvalidOperationException("Expected 1 element");
                }
            }
        }
    }
}
