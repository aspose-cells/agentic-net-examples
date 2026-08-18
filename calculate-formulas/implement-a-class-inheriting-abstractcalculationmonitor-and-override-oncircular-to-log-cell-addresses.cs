// Title: C# Custom AbstractCalculationMonitor to Log Circular Reference Cells in Aspose.Cells
// Description: Shows how to inherit Aspose.Cells.AbstractCalculationMonitor, override the OnCircular method, and write the sheet index and A1‑style address of each circular reference while calculating formulas. The monitor returns true so the engine continues processing.
// Keywords: Aspose.Cells | AbstractCalculationMonitor | OnCircular | circular reference detection | C# .NET | formula calculation | CalculationOptions | log cell addresses | debug Excel formulas | Excel circular reference
// Common Searches: Aspose.Cells custom calculation monitor example | how to capture circular references with AbstractCalculationMonitor | C# log circular reference cells during workbook calculation | override OnCircular in Aspose.Cells | Aspose.Cells CalculationOptions circular reference monitor
// Developer Intent: Create a custom calculation monitor that records the locations of cells involved in circular references when a workbook is evaluated.
// Use Cases: Identify and document circular references in large Excel workbooks without aborting the calculation. | Integrate detailed circular‑reference logging into automated report‑generation pipelines. | Provide developers with sheet, row, and column data to build diagnostic tools for formula errors.
// AI Prompts: Write a C# AbstractCalculationMonitor that writes circular reference details to a text file instead of the console. | Modify the CircularReferenceMonitor to collect cell addresses in a List<string> and return it after workbook.CalculateFormula. | Explain the purpose of the boolean return value from OnCircular and how it affects Aspose.Cells calculation flow.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsCircularDemo
{
    // Custom monitor that logs circular reference cell addresses
    // Shows how to inherit Aspose.Cells.AbstractCalculationMonitor, override the OnCircular method, and write the sheet index and A1‑style address of each circular reference while calculating formulas. The monitor returns true so the engine continues processing.
    public class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected:");
            while (circularCellsData.MoveNext())
            {
                object current = circularCellsData.Current;
                if (current != null)
                {
                    var type = current.GetType();
                    var rowProp = type.GetProperty("Row");
                    var colProp = type.GetProperty("Column");
                    var sheetProp = type.GetProperty("SheetIndex");

                    if (rowProp != null && colProp != null && sheetProp != null)
                    {
                        int row = (int)rowProp.GetValue(current);
                        int col = (int)colProp.GetValue(current);
                        int sheetIdx = (int)sheetProp.GetValue(current);
                        string cellName = CellsHelper.CellIndexToName(row, col);
                        Console.WriteLine($"Sheet {sheetIdx}: {cellName}");
                    }
                    else
                    {
                        // Fallback output if expected properties are missing
                        Console.WriteLine(current);
                    }
                }
            }
            // Return true to let the engine continue calculation for these cells
            return true;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Create a simple circular reference: A1 -> B1, B1 -> A1
                sheet.Cells["A1"].Formula = "=B1";
                sheet.Cells["B1"].Formula = "=A1";

                // Set calculation options with the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor()
                };

                // Perform calculation; the monitor will log circular cells
                workbook.CalculateFormula(options);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Calculation error: {ex.Message}");
            }

            try
            {
                // Save the workbook (optional)
                string outputPath = "CircularReferenceDemo.xlsx";
                Workbook workbook = new Workbook(); // Recreate to include any changes if needed
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Save error: {ex.Message}");
            }
        }
    }
}
