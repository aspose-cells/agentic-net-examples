using System;
using System.Collections;
using Aspose.Cells;

namespace CircularReferenceDetection
{
    class Program
    {
        static void Main()
        {
            // Load the workbook from an existing XLSX file
            Workbook workbook = new Workbook("input.xlsx");

            // Enable calculation chain so that dependents/precedents can be tracked
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Set up calculation options with a custom monitor to catch circular references
            CalculationOptions options = new CalculationOptions();
            options.CalculationMonitor = new CircularReferenceMonitor(workbook);

            // Perform formula calculation; the monitor will be invoked if a circular reference exists
            try
            {
                workbook.CalculateFormula(options);
                Console.WriteLine("Calculation completed without unhandled circular references.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Calculation error: {ex.Message}");
            }

            // Save the workbook after calculation (optional)
            workbook.Save("output.xlsx");
        }

        // Custom monitor derived from AbstractCalculationMonitor to detect circular references
        private class CircularReferenceMonitor : AbstractCalculationMonitor
        {
            private readonly Workbook _workbook;

            public CircularReferenceMonitor(Workbook workbook)
            {
                _workbook = workbook;
            }

            // This method is called when the engine detects a circular reference
            public override bool OnCircular(IEnumerator circularCellsData)
            {
                Console.WriteLine("Circular reference detected!");
                while (circularCellsData.MoveNext())
                {
                    // Each item is a CellArea; display its address
                    if (circularCellsData.Current is CellArea area)
                    {
                        string address = _workbook.Worksheets[0].Cells[area.StartRow, area.StartColumn].Name;
                        Console.WriteLine($" - Cell: {address}");
                    }
                    else
                    {
                        Console.WriteLine($" - Item: {circularCellsData.Current}");
                    }
                }

                // Return true to let the engine continue processing the circular cells,
                // or false to stop further calculation for them.
                return true;
            }
        }
    }
}