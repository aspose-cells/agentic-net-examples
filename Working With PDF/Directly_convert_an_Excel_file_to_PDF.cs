using System;
using Aspose.Cells.Utility;

namespace AsposeCellsExamples
{
    public class ExcelToPdfConverter
    {
        public static void Run(string sourcePath, string pdfPath)
        {
            if (!System.IO.File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                ConversionUtility.Convert(sourcePath, pdfPath);
                Console.WriteLine($"Conversion successful. PDF saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <sourceExcelPath> <outputPdfPath>");
                return;
            }

            string sourcePath = args[0];
            string pdfPath = args[1];
            ExcelToPdfConverter.Run(sourcePath, pdfPath);
        }
    }
}