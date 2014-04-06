﻿using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Excel.Misc
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class CopyContentsTests
    {
        [TestMethod]
        public void TestRowCopyContents()
        {
            var workbook = new XLWorkbook();
            var originalSheet = workbook.Worksheets.Add("original");
            var copyRowSheet = workbook.Worksheets.Add("copy row");
            var copyRowAsRangeSheet = workbook.Worksheets.Add("copy row as range");
            var copyRangeSheet = workbook.Worksheets.Add("copy range");

            originalSheet.Cell("A2").SetValue("test value");
            originalSheet.Range("A2:E2").Merge();

            {
                var originalRange = originalSheet.Range("A2:E2");
                var destinationRange = copyRangeSheet.Range("A2:E2");

                originalRange.CopyTo(destinationRange);
            }
            CopyRowAsRange(originalSheet, 2, copyRowAsRangeSheet, 3);
            {
                var originalRow = originalSheet.Row(2);
                var destinationRow = copyRowSheet.Row(2);
                copyRowSheet.Cell("G2").Value = "must be removed after copy";
                originalRow.CopyTo(destinationRow);
            }
            TestHelper.SaveWorkbook(workbook, @"Misc\CopyRowContents.xlsx");
        }

        private static void CopyRowAsRange(IXLWorksheet originalSheet,  int originalRowNumber, IXLWorksheet destSheet, int destRowNumber)
        {
            {
                var destinationRow = destSheet.Row(destRowNumber);
                destinationRow.Clear();

                var originalRow = originalSheet.Row(originalRowNumber);
                int columnNumber = originalRow.LastCellUsed(true).Address.ColumnNumber;

                var originalRange = originalSheet.Range(originalRowNumber, 1, originalRowNumber, columnNumber);
                var destRange = destSheet.Range(destRowNumber, 1, destRowNumber, columnNumber);
                originalRange.CopyTo(destRange);
            }
        }

        [TestMethod]
        public void CopyConditionalFormatsCount()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.FirstCell().AddConditionalFormat().WhenContains("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.FirstCell();
            Assert.AreEqual(2, ws.ConditionalFormats.Count());
        }

        [TestMethod]
        public void CopyConditionalFormatsRelative()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("=B1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B1" && v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "B2" && v.Value.IsFormula)));
        }

        [TestMethod]
        public void CopyConditionalFormatsFixedStringNum()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("1").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
        }

        [TestMethod]
        public void CopyConditionalFormatsFixedString()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "A";
            ws.Cell("B1").Value = "B";
            ws.Cell("A1").AddConditionalFormat().WhenEquals("A").Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "A" && !v.Value.IsFormula)));
        }

        [TestMethod]
        public void CopyConditionalFormatsFixedNum()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            ws.Cell("A1").Value = "1";
            ws.Cell("B1").Value = "1";
            ws.Cell("A1").AddConditionalFormat().WhenEquals(1).Fill.SetBackgroundColor(XLColor.Blue);
            ws.Cell("A2").Value = ws.Cell("A1");
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
            Assert.IsTrue(ws.ConditionalFormats.Any(cf => cf.Values.Any(v => v.Value.Value == "1" && !v.Value.IsFormula)));
        }
    }
}