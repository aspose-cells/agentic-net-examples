// Title: C# – Load the First Three Worksheets with LightCells in Aspose.Cells
// Description: Demonstrates how to use Aspose.Cells LightCells mode with a custom LightCellsDataHandler to load only the first three worksheets of an XLSX file, dramatically lowering memory consumption. The handler’s StartSheet method returns true for the initial three sheets and false for the rest, allowing partial workbook loading in .NET.
// Keywords: Aspose.Cells LightCells | C# load specific worksheets | partial workbook load | memory optimization Aspose.Cells | LightCellsDataHandler example | .NET Excel streaming | skip worksheets Aspose | GitHub Aspose.Cells sample
// Common Searches: Aspose.Cells load first three sheets | LightCells load selected worksheets .NET | How to limit workbook loading memory usage | LightCellsDataHandler skip sheets example | Partial Excel file load C# Aspose
// Developer Intent: Load only the first three worksheets from an XLSX file using LightCells to minimize memory usage.
// Use Cases: Create a quick summary report by reading just the first three sheets of a massive workbook. | Accelerate data migration scripts that need only the initial worksheets, avoiding unnecessary memory overhead. | Build a web API that processes large Excel uploads but extracts data from the first three sheets to stay within server limits.
// AI Prompts: Generate a LightCellsDataHandler that loads the first N worksheets of a workbook. | Show how to modify the handler to load worksheets based on their names instead of order. | Explain how to retrieve the actual count of worksheets loaded after using LightCells mode.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells LightCells mode with a custom LightCellsDataHandler to load only the first three worksheets of an XLSX file, dramatically lowering memory consumption. The handler’s StartSheet method returns true for the initial three sheets and false for the rest, allowing partial workbook loading in .NET.
class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Create load options and assign a LightCellsDataHandler that limits loading to the first three sheets
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new FirstThreeSheetsHandler();

        // Load the workbook using the LightCells mode
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Display the number of worksheets that were actually loaded
        Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);
        for (int i = 0; i < Math.Min(3, workbook.Worksheets.Count); i++)
        {
            Console.WriteLine($"Sheet {i + 1}: {workbook.Worksheets[i].Name}");
        }
    }

    // LightCellsDataHandler implementation that processes only the first three worksheets
    private class FirstThreeSheetsHandler : LightCellsDataHandler
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
            return false; // Skip remaining sheets
        }

        // The following methods are required by the interface but are not used for sheet filtering
        public bool StartRow(int rowIndex) => true;
        public bool ProcessRow(Row row) => true;
        public bool StartCell(int columnIndex) => true;
        public bool ProcessCell(Cell cell) => true;
    }
}
