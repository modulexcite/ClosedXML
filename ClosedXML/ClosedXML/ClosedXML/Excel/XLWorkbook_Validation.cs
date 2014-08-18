using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace ClosedXML.Excel
{
    public partial class XLWorkbook
    {
        private static IEnumerable<ValidationErrorInfo> GetValidationErrors(SpreadsheetDocument package)
        {
            OpenXmlValidator validator = new OpenXmlValidator();
            if (package == null) return Enumerable.Empty<ValidationErrorInfo>();
            return validator.Validate(package);
        }

        private static bool isValid(SpreadsheetDocument package)
        {
            return GetValidationErrors(package).Count() == 0;
        }

        public static bool isValid(String filePath)
        {
            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                return isValid(package);
            }
        }

        public static IEnumerable<ValidationErrorInfo> GetValidationErrors(String filePath)
        {
            using (var package = SpreadsheetDocument.Open(filePath, false))
            {
                return GetValidationErrors(package);
            }
        }
    }
}