using System;

namespace ClosedXML.Excel
{
    internal class XLRichText : XLFormattedText<IXLRichText>, IXLRichText
    {
        public XLRichText(IXLFontBase defaultFont)
            : base(defaultFont)
        {
            Container = this;
        }

        public XLRichText(XLFormattedText<IXLRichText> defaultRichText, IXLFontBase defaultFont)
            : base(defaultRichText, defaultFont)
        {
            Container = this;
        }

        public XLRichText(String text, IXLFontBase defaultFont)
            : base(text, defaultFont)
        {
            Container = this;
        }
    }
}
