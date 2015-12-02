using System.Collections.Generic;

namespace ClosedXML.Excel
{
    public class XLPicture : IXLPicture
    {
        public string FilePath { get; set; }

        public string Name { get; set; }
        public string Description { get; set; }

        public int WidthPx { get; set; }
        public int HeightPx { get; set; }

        public int ColumnOffsetPx { get; set; }
        public int RowOffsetPx { get; set; }

        public XLPictureType Type { get; set; }

        public bool CanUserChangeAspect { get; set; }
        public bool CanUserCrop { get; set; }
        public bool CanUserMove { get; set; }
        public bool CanUserResize { get; set; }
        public bool CanUserRotate { get; set; }
        public bool CanUserSelect { get; set; }

        public XLPicture()
        {
            Type = XLPictureType.Invalid;
            CanUserChangeAspect = true;
            CanUserCrop = true;
            CanUserMove = true;
            CanUserResize = true;
            CanUserRotate = true;
            CanUserSelect = true;
        }

        public void SetMarker(IXLMarker marker)
        {
            markers = new List<IXLMarker>() { marker };
        }

        public void SetMarker(IXLMarker from, IXLMarker to)
        {
            markers = new List<IXLMarker>() { from, to };
        }

        public IEnumerable<IXLMarker> GetMarkers()
        {
            return (markers != null) ? markers : new List<IXLMarker>();
        }

        public List<IXLMarker> markers;
    }

    public class XLPictures : IXLPictures
    {
        public void Add(IXLPicture picture)
        {
            pictures.Add(picture);
        }

        public IEnumerator<IXLPicture> GetEnumerator()
        {
            return pictures.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private List<IXLPicture> pictures = new List<IXLPicture>();
    }
}
