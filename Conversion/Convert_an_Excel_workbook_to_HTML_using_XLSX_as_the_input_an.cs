using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class XlsxToHtmlConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.html";

            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XlsxToHtmlConverter.Run();
        }
    }
}