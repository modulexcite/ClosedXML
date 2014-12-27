using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ClosedXML.Excel
{
    internal class XLCFNotBlankConverter : IXLCFConverter
    {
        public ConditionalFormattingRule Convert(IXLConditionalFormat cf, int priority, XLWorkbook.SaveContext context)
        {
            var conditionalFormattingRule = new ConditionalFormattingRule { FormatId = (UInt32)context.DifferentialFormats[cf.Style], Type = cf.ConditionalFormatType.ToOpenXml(), Priority = priority };

            var formula = new Formula { Text = "LEN(TRIM(" + cf.Range.RangeAddress.FirstAddress.ToStringRelative(false) + "))>0" };

            conditionalFormattingRule.Append(formula);

            return conditionalFormattingRule;
        }
    }
}
