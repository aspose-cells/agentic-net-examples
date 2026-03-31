using System;
using Aspose.Cells;

namespace AsposeCellsLoadXlsxDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Configure LoadOptions
            LoadOptions loadOptions = new LoadOptions
            {
                // Skip parsing formulas on open to improve performance if formulas are not needed immediately
                ParsingFormulaOnOpen = false,

                // Prefer memory‑efficient loading for large files
                MemorySetting = MemorySetting.MemoryPreference,

                // Keep unparsed data (e.g., comments, drawings) in memory for later use
                KeepUnparsedData = true
            };

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Access the first worksheet to demonstrate that the workbook is loaded
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"Worksheet Name: {firstSheet.Name}");

            // Read a sample cell value (if present)
            Cell sampleCell = firstSheet.Cells["A1"];
            Console.WriteLine($"Cell A1 Value: {sampleCell.Value}");

            // Save the workbook to a new file (preserving the original format)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}