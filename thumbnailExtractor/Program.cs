using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Health.Rise.ClassLibrary;

namespace thumbnailExtractor
{
    public class Program
    {


        static void Main(string[] args)
        {
            try
            {
                string path = (args.Length > 0 && File.Exists(args[0]))
                    ? args[0]
                    : @"..\..\IconFolder\ExcelExample.xlsx";
                int dimensions;
                
                if (args.Length < 2 || !Int32.TryParse(args[1]??string.Empty,out dimensions)) dimensions = 400;
                if (!File.Exists(path))
                {
                    Console.WriteLine("No file found under provided path provided in command line. Please enter valid file path or 'Q' for quit.");
                    do
                    {
                        path = Console.ReadLine();
                    } while (!string.Equals(path, "q", comparisonType: StringComparison.InvariantCultureIgnoreCase)|!File.Exists(path));
                } 
                if (!string.Equals(path, "q",StringComparison.InvariantCultureIgnoreCase))ScreenshotFactory.Screenshot(path, dimensions);
                Console.ReadKey();
            }
            catch (ArgumentNullException e)
            {
                Console.WriteLine(e);
            }
        }
    }
}
