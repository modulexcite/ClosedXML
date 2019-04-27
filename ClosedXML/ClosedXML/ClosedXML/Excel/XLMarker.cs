﻿using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel
{
    public class XLMarker : IXLMarker
    {
        public int ColumnIndex { get; set; }
        public int RowIndex { get; set; }

        public int ColumnOffsetPx { get; set; }
        public int RowOffsetPx { get; set; }
    }

    internal static class XLMarkerConverter
    {
        public static T Convert<T>(IXLMarker marker, Func<int, long> CalcEmuScaleX, Func<int, long> CalcEmuScaleY)
            where T: MarkerType, new()
        {
            return new T()
            {
                ColumnId = new ColumnId((marker.ColumnIndex - 1).ToString()),
                RowId = new RowId((marker.RowIndex - 1).ToString()),
                ColumnOffset = new ColumnOffset(CalcEmuScaleX(marker.ColumnOffsetPx).ToString()),
                RowOffset = new RowOffset(CalcEmuScaleY(marker.RowOffsetPx).ToString()),
            };
        }

        public static XLMarker ConvertBack(MarkerType marker, Func<long, int> CalcPixScaleX, Func<long, int> CalcPixScaleY)
        {
            return new XLMarker()
            {
                ColumnIndex = int.Parse(marker.ColumnId.Text),
                RowIndex = int.Parse(marker.RowId.Text),
                ColumnOffsetPx = CalcPixScaleX(long.Parse(marker.RowOffset.Text)),
                RowOffsetPx = CalcPixScaleY(long.Parse(marker.RowOffset.Text)),
            };
        }
    }
}
