using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Health.Rise.ClassLibrary
{
    /// <summary>
    /// 
    /// </summary>
    /// <seealso cref="thumbnailExtractor.IFileObject" />
    public class ExcelProcessor : FileObject
    {
        #region Fields

        /// <summary>
        /// The excel application
        /// </summary>
        private Application _excelApp;

        #endregion

        #region FileObject Members

        /// <summary>
        /// Takes screenshot from file active area.
        /// </summary>
        /// <param name="path">File path.</param>
        /// <param name="imagePath">Thumbnail output path.</param>
        /// <param name="dimensions">Thumbnail dimensions</param>
        public override void ImageFile(string path, string imagePath, int dimensions)
        {
            try
            {
                Workbook wb;
                Worksheet sheet;
                if (!FileIsKnownTypeAndCorrect(path))
                    Console.WriteLine("Incorrect docx file");
                _excelApp = new Application();
                wb = _excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                sheet = (Worksheet)wb.Sheets[1];

                Range xlRange = sheet.UsedRange;

                xlRange.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap);

                Bitmap image = new Bitmap(ResizeImage(Clipboard.GetImage(), dimensions, dimensions));
                image.Save(imagePath);

                object objFalse = false;
                wb.Close(objFalse, Type.Missing, Type.Missing);
                _excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.InnerException);
            }
            finally
            {
                Marshal.ReleaseComObject(_excelApp);
            }
        }



        #endregion
    }
}
