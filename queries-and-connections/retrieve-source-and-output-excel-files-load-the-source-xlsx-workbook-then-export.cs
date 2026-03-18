using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Define the source XLSX file path and the desired output XLSX file path
            string sourcePath = "source.xlsx";
            string outputPath = "output.xlsx";

            // Load the source workbook using the Workbook(string) constructor
            Workbook workbook = new Workbook(sourcePath);

            // Save the loaded workbook to the output path (XLSX format is inferred from the extension)
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook successfully exported from '{sourcePath}' to '{outputPath}'.");
        }
    }
}