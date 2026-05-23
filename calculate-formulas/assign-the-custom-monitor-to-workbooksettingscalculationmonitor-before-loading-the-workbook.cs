using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMonitorDemo
{
    // Custom monitor that logs calculation progress
    public class MyCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before: Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After : Sheet{sheetIndex}, Row{rowIndex}, Col{columnIndex}");
            Console.WriteLine($"  Original: {OriginalValue}, Calculated: {CalculatedValue}, Changed: {ValueChanged}");
        }

        // Use non‑generic IEnumerator as required by Aspose.Cells
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Continue calculation
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            // Create an instance of the custom monitor
            var monitor = new MyCalculationMonitor();

            // Prepare a workbook (empty by default)
            var workbook = new Workbook();

            // Load an existing workbook file if it exists
            const string inputPath = "Input.xlsx";
            try
            {
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Using a new empty workbook.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            // Attach the monitor via CalculationOptions
            var calcOptions = new CalculationOptions
            {
                CalculationMonitor = monitor
            };

            // Perform formula calculation with monitoring
            try
            {
                workbook.CalculateFormula(calcOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during calculation: {ex.Message}");
                return;
            }

            // Save the workbook after calculation
            const string outputPath = "Output.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}