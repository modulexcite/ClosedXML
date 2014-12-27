using System;
using System.Collections.Generic;

namespace ClosedXML.Excel.CalcEngine
{
    internal static class Is
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("ISBLANK", 1, IsBlank);
        }

        private static object IsBlank(List<Expression> p)
        {
            var v = (string)p[0];
            return String.IsNullOrEmpty(v);
        }
    }
}
