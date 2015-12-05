using DocumentFormat.OpenXml.Packaging;
using System;

namespace ClosedXML.Excel
{
    public enum XLPictureType
    {
        Invalid,

        Bmp,
        Gif,
        Png,
        Tiff,
        Jpeg,
    }

    internal static class XlPictureTypeConverter
    {
        public static ImagePartType Convert(XLPictureType type)
        {
            switch (type)
            {
                case XLPictureType.Bmp: return ImagePartType.Bmp;
                case XLPictureType.Gif: return ImagePartType.Gif;
                case XLPictureType.Png: return ImagePartType.Png;
                case XLPictureType.Tiff: return ImagePartType.Tiff;
                case XLPictureType.Jpeg: return ImagePartType.Jpeg;
                default:
                    throw new ArgumentException("not-supported type: " + type.ToString());
            }
        }

        public static XLPictureType ConvertBack(ImagePart imagePart)
        {
            switch (imagePart.ContentType)
            {
                case "image/bmp": return XLPictureType.Bmp;
                case "image/gif": return XLPictureType.Gif;
                case "image/png": return XLPictureType.Png;
                case "image/jpeg": return XLPictureType.Jpeg;
                case "image/tif": return XLPictureType.Tiff;
                default:
                    throw new ArgumentException("not-supported content-type: " + imagePart.ContentType);
            }
        }
    }
}
