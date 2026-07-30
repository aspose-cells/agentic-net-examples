// Title: C# Custom AbstractCalculationMonitor for Logging Circular Reference Cells in Aspose.Cells
// Description: Demonstrates how to inherit from AbstractCalculationMonitor, override the OnCircular method, and output the A1 addresses of cells that form a circular reference during workbook.CalculateFormula. The sample uses reflection to read row, column, and sheet index from the enumerated cell data, prints each address, and returns true to let the engine continue processing.
// Keywords: Aspose.Cells | AbstractCalculationMonitor | OnCircular override | circular reference logging | C# formula calculation monitor | cell address extraction | reflection in Aspose.Cells | custom calculation monitor example
// Common Searches: how to detect circular references with Aspose.Cells | override OnCircular to log cells in Aspose.Cells .NET | custom calculation monitor for formula evaluation | Aspose.Cells example for circular reference handling | C# log cell addresses when circular reference occurs
// Developer Intent: Create a custom calculation monitor that captures and logs the addresses of cells involved in a circular reference during formula evaluation.
// Use Cases: Debug complex spreadsheets by listing every cell that participates in a circular reference. | Integrate the monitor into an automated testing suite to verify that no unintended circular formulas exist. | Redirect the logged information to a file, database, or monitoring service for audit trails. | Control calculation flow by returning false from OnCircular to abort processing when a circular reference is found.
// AI Prompts: Write a C# AbstractCalculationMonitor that writes circular reference cell addresses to a log file instead of the console. | Show how to attach a custom CircularReferenceMonitor to CalculationOptions and trigger workbook.CalculateFormula to capture circular references. | Explain how modifying OnCircular to return false stops formula calculation after detecting a circular reference.

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;
using System.Reflection;

namespace AsposeCellsCircularDemo
{
    // Demonstrates how to inherit from AbstractCalculationMonitor, override the OnCircular method, and output the A1 addresses of cells that form a circular reference during workbook.CalculateFormula. The sample uses reflection to read row, column, and sheet index from the enumerated cell data, prints each address, and returns true to let the engine continue processing.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Set up a circular reference: A1 -> B1, B1 -> A1
                worksheet.Cells["A1"].Formula = "=B1";
                worksheet.Cells["B1"].Formula = "=A1";

                // Create calculation options and attach the custom monitor
                CalculationOptions options = new CalculationOptions
                {
                    CalculationMonitor = new CircularReferenceMonitor()
                };

                // Perform formula calculation; the monitor will be invoked for circular refs
                workbook.CalculateFormula(options);

                // Save the workbook
                string outputPath = "CircularReferenceDemo.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Custom monitor that logs the addresses of cells involved in a circular reference
    class CircularReferenceMonitor : AbstractCalculationMonitor
    {
        public override bool OnCircular(IEnumerator circularCellsData)
        {
            Console.WriteLine("Circular reference detected. Cells involved:");

            while (circularCellsData.MoveNext())
            {
                object cellObj = circularCellsData.Current;
                if (cellObj != null)
                {
                    // Try to extract Row, Column, and SheetIndex via reflection
                    Type type = cellObj.GetType();
                    PropertyInfo rowProp = type.GetProperty("Row");
                    PropertyInfo colProp = type.GetProperty("Column");
                    PropertyInfo sheetProp = type.GetProperty("SheetIndex");

                    if (rowProp != null && colProp != null && sheetProp != null)
                    {
                        int row = (int)rowProp.GetValue(cellObj);
                        int col = (int)colProp.GetValue(cellObj);
                        int sheetIdx = (int)sheetProp.GetValue(cellObj);
                        string address = CellsHelper.CellIndexToName(row, col);
                        Console.WriteLine($"Sheet {sheetIdx}: {address}");
                    }
                    else
                    {
                        // Fallback: output the raw object if expected properties are missing
                        Console.WriteLine(cellObj.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("null");
                }
            }

            // Return true to allow the engine to continue processing the circular cells
            return true;
        }
    }
}
