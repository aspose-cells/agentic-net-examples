// Title: Aspose.Cells .NET: Custom Calculation Monitor for Real‑Time Formula Evaluation and Circular Reference Detection
// Description: Demonstrates how to create a ProgressCalculationMonitor by extending AbstractCalculationMonitor. The monitor logs before/after each cell calculation, reports original and new values, and captures circular references. It is attached to CalculationOptions and passed to Workbook.CalculateFormula, enabling live tracking of long‑running formula evaluations in a .NET workbook.
// Keywords: Aspose.Cells | C# | AbstractCalculationMonitor | calculation monitor | formula evaluation progress | circular reference detection | Workbook.CalculateFormula | real‑time calculation logging | performance debugging | Excel automation
// Common Searches: Aspose.Cells calculation monitor example | track formula evaluation progress .NET | detect circular references with Aspose.Cells | subscribe to calculation events Aspose.Cells | log cell calculation before and after Aspose.Cells
// Developer Intent: The developer wants to hook into the Aspose.Cells calculation engine to observe each cell’s evaluation, capture circular reference details, and monitor performance during extensive formula processing.
// Use Cases: Log detailed before/after information for every calculated cell to troubleshoot large worksheets. | Identify and list all cells involved in circular references when formulas are evaluated. | Gather metrics such as processed cell count and elapsed time by extending the calculation monitor.
// AI Prompts: Generate a version of AbstractCalculationMonitor that writes progress data to a file instead of the console. | Show how to accumulate timing and cell‑count statistics inside the monitor and display a summary after Workbook.CalculateFormula finishes. | Provide code to pause calculation when a specific cell is reached, wait for user input, then resume using the custom monitor.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace CalculationMonitorDemo
{
    // Custom monitor to track calculation progress
    // Demonstrates how to create a ProgressCalculationMonitor by extending AbstractCalculationMonitor. The monitor logs before/after each cell calculation, reports original and new values, and captures circular references. It is attached to CalculationOptions and passed to Workbook.CalculateFormula, enabling live tracking of long‑running formula evaluations in a .NET workbook.
    public class ProgressCalculationMonitor : AbstractCalculationMonitor
    {
        // Called before each cell is calculated
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"[Before] Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
        }

        // Called after each cell is calculated
        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"[After]  Sheet {sheetIndex}, Row {rowIndex}, Column {columnIndex}");
            Console.WriteLine($"    Original: {OriginalValue}, New: {CalculatedValue}, Changed: {ValueChanged}");
        }

        // Called when a circular reference is detected
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                // Each item provides Row, Column and Worksheet properties
                var cell = circularCellsData.Current;
                if (cell != null)
                {
                    try
                    {
                        dynamic d = cell; // use dynamic to access properties at runtime
                        int row = d.Row;
                        int col = d.Column;
                        Worksheet ws = d.Worksheet;
                        string address = CellsHelper.CellIndexToName(row, col);
                        Console.WriteLine($"    {ws.Name}!{address}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"    Unable to read cell info: {ex.Message}");
                    }
                }
            }
            // Continue calculation for circular cells
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a range with values
                for (int r = 0; r < 10; r++)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        sheet.Cells[r, c].PutValue(r + c);
                    }
                }

                // Add some formulas that depend on the populated range
                sheet.Cells["K1"].Formula = "=SUM(A1:J10)";          // Large sum
                sheet.Cells["K2"].Formula = "=AVERAGE(A1:J10)";     // Average
                sheet.Cells["K3"].Formula = "=MAX(A1:J10)";         // Max
                sheet.Cells["K4"].Formula = "=MIN(A1:J10)";         // Min

                // Introduce a circular reference for demonstration
                sheet.Cells["L1"].Formula = "=L2";
                sheet.Cells["L2"].Formula = "=L1";

                // Set up calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new ProgressCalculationMonitor(),
                    Recursive = true,
                    IgnoreError = false
                };

                // Perform calculation with monitoring
                workbook.CalculateFormula(options);

                // Output some results after calculation
                Console.WriteLine("\n--- Calculation Results ---");
                Console.WriteLine($"K1 (SUM): {sheet.Cells["K1"].Value}");
                Console.WriteLine($"K2 (AVERAGE): {sheet.Cells["K2"].Value}");
                Console.WriteLine($"K3 (MAX): {sheet.Cells["K3"].Value}");
                Console.WriteLine($"K4 (MIN): {sheet.Cells["K4"].Value}");

                // Save the workbook (optional)
                string outputPath = "CalculationMonitorDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"\nWorkbook saved to: {Path.GetFullPath(outputPath)}");
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
}
