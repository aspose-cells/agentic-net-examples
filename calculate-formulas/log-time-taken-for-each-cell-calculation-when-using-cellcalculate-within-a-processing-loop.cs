// Title: Log execution time for each cell’s formula calculation using Cell.Calculate in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to measure time of Cell.Calculate for each cell | log per‑cell formula evaluation duration with Stopwatch in Aspose.Cells | benchmark individual cell calculations in a .NET Excel workbook | track performance of Cell.Calculate across a range of cells using Aspose.Cells | measure execution time of Excel cell formulas in C# with Aspose.Cells
// Tags: cell calculation timing Aspose.Cells | per‑cell formula performance measurement .NET | stopwatch based Excel cell evaluation logging | benchmark individual cell calculations C# | Aspose.Cells performance logging for formulas

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, defines a cell range (A1:D10), and iterates through each cell. For every cell it starts a Stopwatch, calls Cell.Calculate with default CalculationOptions, stops the timer, and writes the elapsed milliseconds to the console. After processing the range, the workbook is saved to a new file, providing a simple way to log per‑cell formula calculation times.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to process (example: A1 to D10)
            CellArea range = CellArea.CreateCellArea("A1", "D10");

            // Prepare calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // Loop through each cell in the defined range
            for (int row = range.StartRow; row <= range.EndRow; row++)
            {
                for (int col = range.StartColumn; col <= range.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Start timing
                    Stopwatch sw = Stopwatch.StartNew();

                    // Perform calculation for the current cell with options
                    cell.Calculate(calcOptions);

                    // Stop timing
                    sw.Stop();

                    // Log the time taken
                    Console.WriteLine($"Cell {cell.Name} calculated in {sw.Elapsed.TotalMilliseconds} ms");
                }
            }

            // Save the workbook after processing (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
