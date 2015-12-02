using System;
using ClosedXML.Excel;
using System.IO;

namespace ClosedXML_Examples.Pictures
{
    public class AddingPicture
    {
        public void Create(String filePath)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sample Sheet");

            var filepath = Path.GetFullPath(
                Path.Combine("..", "..", "Pictures", "test.jpg"));

            var picture = new XLPicture()
            {
                FilePath = filepath,
                WidthPx = 100,
                HeightPx = 100,
                Type = XLPictureType.Jpeg,
                CanUserChangeAspect = false,
            };
            picture.SetMarker(new XLMarker()
            {
                RowIndex = 12,
                ColumnIndex = 13,
                RowOffsetPx = 5,
                ColumnOffsetPx = 10,
            });
            worksheet.Pictures.Add(picture);

            var picture1 = new XLPicture()
            {
                FilePath = filepath,
                Type = XLPictureType.Jpeg,
            };
            picture1.SetMarker(
                new XLMarker()
                {
                    RowIndex = 14,
                    ColumnIndex = 15,
                    RowOffsetPx = 5,
                    ColumnOffsetPx = 10,
                },
                new XLMarker()
                {
                    RowIndex = 18,
                    ColumnIndex = 19,
                });
            worksheet.Pictures.Add(picture1);

            var picture2 = new XLPicture()
            {
                FilePath = filepath,
                Type = XLPictureType.Jpeg,
            };
            worksheet.Pictures.Add(picture2);

            workbook.SaveAs(filePath);
        }
    }
}
