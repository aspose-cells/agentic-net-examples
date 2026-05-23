using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularReferenceDemo
{
    // Custom monitor to handle circular references and report detailed cell positions
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        // Called when the calculation engine detects a circular reference
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Involved cells:");

            // Enumerate all CalculationCell objects that are part of the circular chain
            while (circularCellsData.MoveNext())
            {
                if (circularCellsData.Current is CalculationCell calcCell)
                {
                    int sheetIndex = calcCell.Worksheet.Index;
                    string sheetName = calcCell.Worksheet.Name;
                    int rowIndex = calcCell.CellRow;      // zero‑based row index
                    int colIndex = calcCell.CellColumn;   // zero‑based column index

                    // Convert column index to Excel column letter for readability
                    string columnLetter = GetColumnLetter(colIndex);

                    Console.WriteLine(
                        $"  Sheet {sheetIndex} ('{sheetName}'): Cell {columnLetter}{rowIndex + 1} (Row={rowIndex}, Column={colIndex})");
                }
            }

            // Return true to let the engine continue calculating these cells,
            // or false to mark them as calculated without further processing.
            return true;
        }

        // Helper to convert zero‑based column index to Excel column letters (e.g., 0 -> A)
        private static string GetColumnLetter(int columnIndex)
        {
            const int lettersInAlphabet = 26;
            string columnLetter = string.Empty;
            int dividend = columnIndex + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % lettersInAlphabet;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                dividend = (dividend - modulo) / lettersInAlphabet;
            }

            return columnLetter;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Set up a simple circular reference scenario:
                // A1 depends on B1, and B1 depends on A1
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Configure calculation options to use the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor()
                    // Other options retain their default values
                };

                // Perform formula calculation; the monitor will be invoked automatically
                workbook.CalculateFormula(options);

                // Save the workbook (optional, just to demonstrate full lifecycle)
                string outputPath = "CircularReferenceDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}