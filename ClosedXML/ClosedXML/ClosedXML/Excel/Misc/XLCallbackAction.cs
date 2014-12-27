using System;

namespace ClosedXML.Excel.Misc
{
    internal class XLCallbackAction
    {
        public XLCallbackAction(Action<XLRange, Int32> action)
        {
            this.Action = action;
        }

        public Action<XLRange, Int32> Action { get; set; }
    }
}
