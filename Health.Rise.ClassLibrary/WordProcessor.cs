using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Media;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Media.Imaging;
using Application = Microsoft.Office.Interop.Word.Application;
using DataFormats = System.Windows.Forms.DataFormats;

namespace Health.Rise.ClassLibrary
{
    public class WordProcessor: FileObject
    {
        private Application _wordApplication;
        public override void ImageFile(string path, string imagePath, int dimensions)
        {
            try
            {
                _wordApplication = new Application();
                var doc = _wordApplication.Documents.Open(path);
                var page = doc.ActiveWindow.ActivePane.Pages;
                using (var stream = new MemoryStream((byte[])page[1].EnhMetaFileBits))
                {
                    if (!FileIsKnownTypeAndCorrect(path))
                        Console.WriteLine("Incorrect docx file");

                    var image = ResizeImage(Image.FromStream(stream), dimensions, dimensions);
                    var pngTarget = imagePath;
                    image.Save(pngTarget, ImageFormat.Png);
                }
                doc.Close();
                _wordApplication.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.InnerException);
            }
            finally
            {
                Marshal.ReleaseComObject(_wordApplication);
            }
        }




    }
}
