using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel spreadsheet (Aspose.Cells does not support .numbers files)
            string sourceFile = "input.xlsx";

            // Desired PDF output path
            string pdfFile = "output.pdf";

            // If the source file does not exist, create a simple workbook for demonstration
            if (!File.Exists(sourceFile))
            {
                var tempWb = new Workbook();
                var sheet = tempWb.Worksheets[0];
                sheet.Name = "Demo";
                sheet.Cells["A1"].PutValue("Sample data");
                tempWb.Save(sourceFile);
            }

            // Load the Excel file into a Workbook instance
            Workbook workbook = new Workbook(sourceFile);

            // Save the workbook as PDF
            workbook.Save(pdfFile, SaveFormat.Pdf);

            Console.WriteLine($"Conversion completed: '{sourceFile}' → '{pdfFile}'");
        }
    }
}