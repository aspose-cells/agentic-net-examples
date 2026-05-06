using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source Excel file (must contain embedded Office Add‑Ins)
            string sourceFileName = "input_with_addins.xlsx";
            string sourcePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, sourceFileName);

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Desired output PDF file path
            string destFileName = "output.pdf";
            string destPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, destFileName);

            // Load the workbook and save it as PDF.
            var workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine($"Excel workbook successfully converted to PDF: {destPath}");
        }
    }
}