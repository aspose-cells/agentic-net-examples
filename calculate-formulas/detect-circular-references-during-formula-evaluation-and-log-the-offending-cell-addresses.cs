using System;
using System.Collections;
using Aspose.Cells;

class CircularReferenceLogger : AbstractCalculationMonitor
{
    // Called when a circular reference is detected.
    // The enumerator contains Cell objects that are part of the circular chain.
    public override bool OnCircular(IEnumerator circularCellsData)
    {
        Console.WriteLine("Circular reference detected. Offending cells:");

        while (circularCellsData.MoveNext())
        {
            // Each item is a Cell; retrieve its address and sheet name.
            var cell = circularCellsData.Current as Cell;
            if (cell != null)
            {
                string address = CellsHelper.CellIndexToName(cell.Row, cell.Column);
                string sheetName = cell.Worksheet?.Name ?? "Unknown Sheet";
                Console.WriteLine($"{sheetName}: {address}");
            }
            else
            {
                // Fallback if casting fails.
                Console.WriteLine(circularCellsData.Current?.ToString());
            }
        }

        // Return true to let the engine continue processing the circular cells,
        // or false to stop further calculation for them.
        return true;
    }
}

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a simple circular reference: A1 -> B1 -> A1.
            sheet.Cells["A1"].Formula = "=B1";
            sheet.Cells["B1"].Formula = "=A1";

            // Configure calculation options to use our custom monitor.
            CalculationOptions options = new CalculationOptions
            {
                CalculationMonitor = new CircularReferenceLogger()
            };

            // Perform formula calculation; the monitor will be invoked on circular detection.
            workbook.CalculateFormula(options);

            // Save the workbook (the file will contain the formulas as set above).
            string outputPath = "CircularReferenceDemo.xlsx";

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}