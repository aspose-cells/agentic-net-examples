using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaClearDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path) or create a new one.
            // Here we demonstrate both scenarios.
            Workbook workbook;
            string inputPath = "InputWorkbook.xlsx"; // Path to the workbook that may have a print area defined.

            if (System.IO.File.Exists(inputPath))
            {
                // Load the workbook from file.
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook if the file does not exist.
                workbook = new Workbook();
                // Example: define a print area to later clear it.
                workbook.Worksheets[0].PageSetup.PrintArea = "A1:C10";
            }

            // Iterate through all worksheets and clear any defined print area.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Setting PrintArea to an empty string removes the previously defined range.
                sheet.PageSetup.PrintArea = string.Empty;
            }

            // Save the workbook after clearing the print areas.
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Print areas cleared and workbook saved to '{outputPath}'.");
        }
    }
}