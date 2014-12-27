namespace ClosedXML.Excel
{
    internal class XLRangeParameters
    {
        #region Constructor

        public XLRangeParameters(XLRangeAddress rangeAddress, IXLStyle defaultStyle)
        {
            RangeAddress = rangeAddress;
            _ignoreEvents = !rangeAddress.Worksheet.EventTrackingEnabled;
            DefaultStyle = defaultStyle;
        }

        #endregion Constructor

        #region Properties

        public XLRangeAddress RangeAddress { get; private set; }

        public IXLStyle DefaultStyle { get; private set; }

        private bool _ignoreEvents;

        public bool IgnoreEvents
        {
            get { return _ignoreEvents; }
            set
            {
                _ignoreEvents = value;
            }
        }

        #endregion Properties
    }
}
