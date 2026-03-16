using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class ExcelToHtmlConverter
    {
        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Desired output HTML file path
            string outputPath = "output.html";

            // Convert the XLSX workbook to HTML using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{outputPath}'");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToHtmlConverter.Run();
        }
    }
}