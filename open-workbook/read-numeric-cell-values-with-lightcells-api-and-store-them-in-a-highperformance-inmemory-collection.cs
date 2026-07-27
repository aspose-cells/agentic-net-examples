// Title: Read numeric cells from a large Excel file with Aspose.Cells LightCells and store them in a high‑performance List<double>
// Description: Loads a workbook in LightCells streaming mode, uses a custom LightCellsDataHandler to scan every sheet, row and cell, captures each numeric value via DoubleValue, adds it to a List<double>, returns false to keep the cell out of the workbook model for minimal memory usage, and saves the workbook unchanged.
// Keywords: Aspose.Cells LightCells | C# read numeric cells | streaming Excel processing | memory‑efficient Excel read | .NET high‑performance collection | List<double> Excel values | LightCellsDataHandler example | large workbook numeric extraction
// Common Searches: Aspose.Cells LightCells read only numbers | C# extract numeric values from big Excel file | How to use LightCellsDataHandler to collect doubles | Memory‑saving Excel read with Aspose.Cells | Stream large workbook and get numeric cells
// Developer Intent: Efficiently retrieve all numeric values from a massive Excel workbook without loading the full object model.
// Use Cases: Perform statistical analysis on numeric data extracted from a multi‑gigabyte spreadsheet while keeping RAM usage low. | Feed collected numbers into a database, analytics pipeline, or machine‑learning model after streaming read. | Validate ranges, detect outliers, or apply business rules to numeric cells during processing before saving the file.
// AI Prompts: Create a LightCellsDataHandler that captures numeric values and also logs each cell's address. | Adapt the handler to filter numbers above a configurable threshold and store them in a thread‑safe ConcurrentBag<double>. | Show how to serialize the List<double> of extracted values to JSON after the workbook has been processed.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads a workbook in LightCells streaming mode, uses a custom LightCellsDataHandler to scan every sheet, row and cell, captures each numeric value via DoubleValue, adds it to a List<double>, returns false to keep the cell out of the workbook model for minimal memory usage, and saves the workbook unchanged.
class Program
{
    static void Main()
    {
        // Input workbook path (can be any large Excel file)
        string inputPath = "LargeData.xlsx";
        // Output path – the workbook is saved unchanged after processing
        string outputPath = "Processed.xlsx";

        // Configure LoadOptions to use LightCells mode with a custom handler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new NumericValuesHandler();

        // Load the workbook in streaming (LightCells) mode
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Retrieve the collected numeric values after processing
        var numericValues = ((NumericValuesHandler)loadOptions.LightCellsDataHandler).NumericValues;
        Console.WriteLine($"Collected {numericValues.Count} numeric values.");

        // Save the workbook (unchanged) – demonstrates use of the save lifecycle rule
        workbook.Save(outputPath);
    }

    // Custom LightCellsDataHandler that extracts numeric cell values
    private class NumericValuesHandler : LightCellsDataHandler
    {
        // High‑performance in‑memory collection for numeric values
        public readonly List<double> NumericValues = new List<double>();

        // Process every sheet
        public bool StartSheet(Worksheet sheet) => true;

        // Process every row
        public bool StartRow(int rowIndex) => true;

        // Process every cell in a row
        public bool StartCell(int columnIndex) => true;

        // No special row processing needed
        public bool ProcessRow(Row row) => true;

        // Called for each cell; store numeric values and release the cell from memory
        public bool ProcessCell(Cell cell)
        {
            if (cell.IsNumericValue)
            {
                // DoubleValue works for int, double, and datetime numeric types
                NumericValues.Add(cell.DoubleValue);
            }
            // Return false to keep the cell out of the workbook model and save memory
            return false;
        }
    }
}
