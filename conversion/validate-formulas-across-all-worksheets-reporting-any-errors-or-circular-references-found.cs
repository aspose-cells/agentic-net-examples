using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace FormulaValidationDemo
{
    // Custom monitor to capture circular reference information
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public List<string> CircularCells { get; } = new List<string>();

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            // Called when a circular reference is detected
            while (circularCellsData.MoveNext())
            {
                // Each item is a CalculationCell representing a cell in the circular chain
                if (circularCellsData.Current is CalculationCell cc)
                {
                    // Use dynamic to access Row/Column at runtime (API may expose different property names)
                    dynamic d = cc;
                    int row = d.Row;      // zero‑based row index
                    int column = d.Column; // zero‑based column index

                    string cellName = $"{cc.Worksheet.Name}!{row + 1}:{column + 1}";
                    CircularCells.Add(cellName);
                }
            }
            // Continue normal calculation after reporting
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook --------------------
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // Sample data
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(0);

                // Valid formula
                cells["B1"].Formula = "=A1*2";

                // Formula that will cause a division by zero error
                cells["B2"].Formula = "=A1/A2";

                // Circular reference example
                cells["C1"].Formula = "=C2+1";
                cells["C2"].Formula = "=C1+1";

                // -------------------- Set calculation options --------------------
                var monitor = new CircularReferenceMonitor();

                CalculationOptions options = new CalculationOptions
                {
                    // Do not throw on formula errors; they will be marked in cells
                    IgnoreError = true,
                    // Attach monitor to capture circular references
                    CalculationMonitor = monitor
                };

                // -------------------- Perform calculation --------------------
                wb.CalculateFormula(options);

                // -------------------- Report formula errors --------------------
                Console.WriteLine("Formula Errors:");
                foreach (Worksheet sheet in wb.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        // Check only cells that contain formulas
                        if (cell.IsFormula && cell.Type == CellValueType.IsError)
                        {
                            Console.WriteLine($"{sheet.Name}!{cell.Name} -> Error: {cell.StringValue}");
                        }
                    }
                }

                // -------------------- Report circular references --------------------
                Console.WriteLine("\nCircular References Detected:");
                if (monitor.CircularCells.Count == 0)
                {
                    Console.WriteLine("None");
                }
                else
                {
                    foreach (string cellRef in monitor.CircularCells)
                    {
                        Console.WriteLine(cellRef);
                    }
                }

                // -------------------- Save workbook (optional) --------------------
                wb.Save("FormulaValidationResult.xlsx");
            }
            catch (Exception ex)
            {
                // Runtime safety: log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}