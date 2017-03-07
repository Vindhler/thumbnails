using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace thumbnailExtractor
{
    public class Program
    {
        /// <summary>
        /// Image file resize to thumbnail dimensions.
        /// </summary>
        /// <param name="path">Input image path.</param>
        /// <param name="outputPath">The output image path.</param>
        public void ImageFileResize(string path, string outputPath)
        {
            Image image = Image.FromFile(path);
            Image result = image.GetThumbnailImage(400, 400, null, IntPtr.Zero);
            result.Save(outputPath);
        }


        [STAThread]
        static void Main(string[] args)
        {
            IFileObject file = new ExcelProcessor();
            file.ImageFile(@"E:\ExcelExample.xlsx", @"E:\screenshotexc.png");
            Console.ReadKey();
        }
    }
}
