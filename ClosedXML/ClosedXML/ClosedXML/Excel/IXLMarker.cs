
namespace ClosedXML.Excel
{
    public interface IXLMarker
    {
        int ColumnIndex { get; set; }
        int RowIndex { get; set; }

        int ColumnOffsetPx { get; set; }
        int RowOffsetPx { get; set; }
    }
}
