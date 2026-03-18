using System;
using Aspose.Cells;

namespace ExportWorkbook
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (any supported format)
            string sourcePath = "input.xlsx"; // TODO: set actual source file path

            // Load the workbook – this retains all worksheets, formulas, formatting, and metadata
            Workbook workbook = new Workbook(sourcePath);

            // Define the output XLSX file path
            string outputPath = "output.xlsx";

            // Save the workbook as XLSX, preserving everything
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Clean up resources
            workbook.Dispose();

            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
    }
}