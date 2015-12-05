using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.IO;

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

        internal ImagePart ImagePart { get; private set; }

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

        public Stream OpenImageStream(FileMode mode = FileMode.Open, FileAccess access = FileAccess.Read)
        {
            return (ImagePart != null) ? ImagePart.GetStream(mode, access) : null;
        }

        internal void SetImagePart(ImagePart part)
        {
            ImagePart = part;
        }

        private List<IXLMarker> markers;
    }

    public class XLPictures : IXLPictures
    {
        public void Add(IXLPicture picture)
        {
            pictures.Add(picture);
        }

        public void AddRange(IEnumerable<IXLPicture> pictures)
        {
            this.pictures.AddRange(pictures);
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
