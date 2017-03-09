using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Health.Rise.ClassLibrary
{
    public static class ScreenshotFactory
    {

        public static bool IsWord(this string filename)
        {
            return Path.GetExtension(filename).Contains("doc"); 
        }

        public static bool IsExcel(this string filename)
        {
            return Path.GetExtension(filename).Contains("xls");
        }


        public static void Screenshot(string filename = @"..\..\IconFolder\ExcelExample.xlsx", int dimensions = 400)
        {
            try
            {
                string thumbNailFileName =
                    $"{Path.GetDirectoryName(filename)}{Path.GetFileNameWithoutExtension(filename)}-thumbnail-{Guid.NewGuid().ToString().Substring(0, 4)}.png";

                FileObject file = null;
                if (filename.IsWord() == filename.IsExcel()) Console.WriteLine("file  not recognized correctly");
                else if (filename.IsWord())
                {
                    file = new WordProcessor();
                    file.ImageFile(filename, thumbNailFileName, dimensions);
                }
                else if (filename.IsExcel())
                {
                    file = new ExcelProcessor();
                    file.ImageFile(filename, thumbNailFileName, dimensions);
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File loading failed with message: {ex.Message}");

            }

        }
    }
}
