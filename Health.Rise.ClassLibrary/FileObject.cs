using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Health.Rise.ClassLibrary
{
    public abstract class FileObject
    {
 
        public abstract void ImageFile(string path, string imagePath, int dimensions);

        public bool FileIsKnownTypeAndCorrect(string path)
        {
            using (var stream = File.OpenRead(path))
            {
                string ext = Path.GetExtension(path);
                string format = FileDeterminator.DetermineFormat(stream);
                if (string.Equals(format, "DOCorXLS") & (ext.Contains("doc") | (ext.Contains("xls")))) return true;
                if (string.Equals(format, "DOCXorXLSX") & (ext.Contains("docx") | (ext.Contains("xlsx")))) return true;
                if (string.Equals(format, "") & (ext.Contains("docx") | (ext.Contains("xlsx")))) return true;
                return false;
            }
        }

        public static Bitmap ResizeImage(Image image, int width, int height)
        {
            var destRect = new Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }

            return destImage;
        }

       
    }
}