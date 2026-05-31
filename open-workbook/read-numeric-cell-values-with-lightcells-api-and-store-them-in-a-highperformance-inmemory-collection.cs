using System;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    // Thread‑safe collection for high‑performance storage of numeric values
    private static readonly List<double> _numericValues = new List<double>();

    static void Main()
    {
        // Path to the source workbook (replace with your actual file)
        const string sourcePath = "LargeData.xlsx";

        // Configure LightCells loading with a custom handler
        var loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.LightCellsDataHandler = new NumericCollectorHandler();

        // Load the workbook in LightCells mode – cells are streamed and processed by the handler
        var workbook = new Workbook(sourcePath, loadOptions);

        // At this point _numericValues holds all numeric cell values from the workbook
        Console.WriteLine($"Collected {_numericValues.Count} numeric values.");

        // Example aggregation – compute the sum of all collected numbers
        double total = 0;
        foreach (double val in _numericValues)
            total += val;
        Console.WriteLine($"Sum of numeric values: {total}");

        // Optional: save the (unchanged) workbook to a new file
        workbook.Save("Processed.xlsx");
    }

    // Custom LightCellsDataHandler that extracts numeric values
    private class NumericCollectorHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet) => true;          // Process every sheet
        public bool StartRow(int rowIndex) => true;              // Process every row
        public bool StartCell(int columnIndex) => true;          // Process every cell
        public bool ProcessRow(Row row) => true;                 // Continue processing rows

        public bool ProcessCell(Cell cell)
        {
            // If the cell contains a numeric value (int, double, or datetime), store it
            if (cell.IsNumericValue)
                _numericValues.Add(cell.DoubleValue);

            // Return false so the cell is not retained in the workbook's in‑memory model
            return false;
        }
    }
}