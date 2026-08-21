// Title: Attach a Custom AbstractCalculationMonitor to an Aspose.Cells Workbook for Formula Calculation (C#)
// Description: Demonstrates how to create a class that inherits from AbstractCalculationMonitor, assign it to Workbook.Settings.CalculationMonitor (or via CalculationOptions), run workbook.CalculateFormula, capture before/after callbacks, detect circular references, and save the updated file.
// Keywords: Aspose.Cells | C# | AbstractCalculationMonitor | custom calculation monitor | formula calculation | circular reference detection | Workbook.CalculateFormula | calculation callbacks | Excel automation | performance logging
// Common Searches: Aspose.Cells custom calculation monitor example | How to use AbstractCalculationMonitor in C# | Assign calculation monitor before workbook.CalculateFormula | Log cell calculation events with Aspose.Cells | Detect circular references using Aspose.Cells monitor
// Developer Intent: Add a custom calculation monitor to a workbook and execute formula evaluation with real‑time callbacks.
// Use Cases: Trace the start and end of each cell's calculation to diagnose performance bottlenecks. | Log original and new cell values after formula evaluation for audit trails. | Identify and handle circular references during spreadsheet processing.
// AI Prompts: Generate C# code that sets Workbook.Settings.CalculationMonitor to a custom monitor before loading a workbook with Aspose.Cells. | Show how to redirect the monitor's console output to a file or logging framework. | Explain how to parse the IEnumerator passed to OnCircular to list the cells involved in a circular reference.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomMonitorDemo
{
    // Custom monitor that inherits from AbstractCalculationMonitor
    // Demonstrates how to create a class that inherits from AbstractCalculationMonitor, assign it to Workbook.Settings.CalculationMonitor (or via CalculationOptions), run workbook.CalculateFormula, capture before/after callbacks, detect circular references, and save the updated file.
    public class MyCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"Before calculating Sheet{sheetIndex}, Row{rowIndex}, Column{columnIndex}");
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            Console.WriteLine($"After calculating Sheet{sheetIndex}, Row{rowIndex}, Column{columnIndex}");
            Console.WriteLine($"Original: {OriginalValue}, New: {CalculatedValue}, Changed: {ValueChanged}");
        }

        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected.");
            // Continue calculation despite circular reference
            return true;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create an instance of the custom calculation monitor
                MyCalculationMonitor monitor = new MyCalculationMonitor();

                Workbook workbook;

                // Ensure the input file exists; if not, create a simple workbook for demonstration
                const string inputPath = "Input.xlsx";
                if (!File.Exists(inputPath))
                {
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue(10);
                    sheet.Cells["A2"].PutValue(20);
                    sheet.Cells["A3"].Formula = "=A1+A2";
                    workbook.Save(inputPath);
                }
                else
                {
                    // Load the existing workbook
                    workbook = new Workbook(inputPath);
                }

                // Prepare calculation options and attach the monitor
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = monitor
                };

                // Perform formula calculation with the custom monitor active
                workbook.CalculateFormula(calcOptions);

                // Save the workbook after calculation
                const string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Calculation completed. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
