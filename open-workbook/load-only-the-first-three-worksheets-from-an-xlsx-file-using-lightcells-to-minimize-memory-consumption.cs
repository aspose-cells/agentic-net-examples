// Title: Memory‑Efficiently Load the First Three Worksheets with LightCells in Aspose.Cells for .NET
// Description: Demonstrates how to create a custom LightCellsDataHandler that processes only the first three worksheets, configure LoadOptions for LightCells mode, load a large XLSX file partially, enumerate the loaded sheets, and save the trimmed workbook.
// Keywords: Aspose.Cells LightCells | load specific worksheets .NET | partial workbook loading | memory efficient Excel processing | C# LightCellsDataHandler | skip worksheets Aspose | large workbook performance
// Common Searches: Aspose.Cells load first three sheets | LightCells load selected worksheets | C# memory efficient Excel loading | How to skip worksheets with LightCells | Partial workbook load Aspose.Cells
// Developer Intent: Load only the first three worksheets of an XLSX file to minimize memory consumption.
// Use Cases: Extract data from the initial three sheets of a massive workbook without loading the entire file. | Create a lightweight copy of a multi‑sheet workbook that contains only the first three worksheets for downstream analysis. | Generate a quick summary report by processing just the first three sheets, reducing runtime and memory overhead.
// AI Prompts: Show how to modify the LightCellsDataHandler to load a configurable number of worksheets based on a parameter. | Provide an example of loading worksheets by name using LightCells instead of by index. | Explain how to combine LightCells with streaming to export the first three sheets to CSV files efficiently.

using System;
using Aspose.Cells;

namespace LightCellsFirstThreeSheetsDemo
{
    // Custom LightCellsDataHandler that processes only the first three worksheets
    // Demonstrates how to create a custom LightCellsDataHandler that processes only the first three worksheets, configure LoadOptions for LightCells mode, load a large XLSX file partially, enumerate the loaded sheets, and save the trimmed workbook.
    public class FirstThreeSheetsHandler : LightCellsDataHandler
    {
        private int _processedSheets = 0;

        // Called before reading each worksheet; return true only for the first three sheets
        public bool StartSheet(Worksheet sheet)
        {
            if (_processedSheets < 3)
            {
                _processedSheets++;
                return true; // Process this sheet
            }
            return false; // Skip this sheet
        }

        // Process all rows
        public bool StartRow(int rowIndex) => true;

        // Process all rows
        public bool ProcessRow(Row row) => true;

        // Process all cells
        public bool StartCell(int columnIndex) => true;

        // Process all cells
        public bool ProcessCell(Cell cell) => true;
    }

    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string sourcePath = "LargeWorkbook.xlsx";

            // Configure LoadOptions to use LightCells mode with the custom handler
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LightCellsDataHandler = new FirstThreeSheetsHandler();

            // Load the workbook; only the first three worksheets will be loaded into memory
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Verify the number of loaded worksheets (should be <= 3)
            Console.WriteLine($"Worksheets loaded: {workbook.Worksheets.Count}");

            // Example: iterate through the loaded sheets and print their names
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Sheet: {sheet.Name}");
            }

            // Save the partially loaded workbook if needed
            string outputPath = "FirstThreeSheetsOnly.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
