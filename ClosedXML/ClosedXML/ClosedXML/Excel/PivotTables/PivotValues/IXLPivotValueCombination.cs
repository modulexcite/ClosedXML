using System;

namespace ClosedXML.Excel
{
    public interface IXLPivotValueCombination
    {
        IXLPivotValue And(String item);

        IXLPivotValue AndPrevious();

        IXLPivotValue AndNext();
    }
}
