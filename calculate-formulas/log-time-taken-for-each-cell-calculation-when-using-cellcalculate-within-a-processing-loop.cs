// Title: Measure and log the execution time of each cell’s formula calculation with Cell.Calculate in Aspose.Cells for .NET
// Common Searches: how to log execution time of individual cell formulas using Aspose.Cells C# | benchmark per‑cell calculation performance with Cell.Calculate in .NET | record milliseconds taken by each Excel cell calculation in Aspose.Cells | use Stopwatch to time Cell.Calculate calls in a loop Aspose.Cells | profile worksheet cell formula evaluation duration Aspose.Cells .NET
// Tags: cell.Calculate performance timing .NET | per‑cell formula execution logging Aspose.Cells | stopwatch measurement of Excel cell calculations | benchmark individual worksheet cells C# | log calculation latency for Excel formulas Aspose

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, iterates over cells A1:D10 on the first worksheet, and uses a Stopwatch to capture the milliseconds each Cell.Calculate call takes (with default CalculationOptions). It writes the timing for every cell to the console and saves the modified workbook.
class CellCalculationTimer
{
    static void Main()
    {
        try
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to process (A1:D10)
            CellArea area = CellArea.CreateCellArea("A1", "D10");

            // Prepare calculation options (default settings)
            CalculationOptions calcOptions = new CalculationOptions();

            // Iterate through each cell in the range and time its calculation
            for (int row = area.StartRow; row <= area.EndRow; row++)
            {
                for (int col = area.StartColumn; col <= area.EndColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    Stopwatch sw = Stopwatch.StartNew();
                    cell.Calculate(calcOptions); // Perform calculation with options
                    sw.Stop();

                    Console.WriteLine($"Cell {cell.Name} calculated in {sw.Elapsed.TotalMilliseconds} ms");
                }
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
