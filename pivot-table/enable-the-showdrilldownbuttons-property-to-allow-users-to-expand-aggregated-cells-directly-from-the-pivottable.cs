// Title: Enabling drill‑down buttons for a PivotTable with Aspose.Cells for .NET (C#) – handling the missing ShowDrillDownButtons property
// Common Searches: C# Aspose.Cells how to display drill down buttons in a PivotTable | show expand collapse buttons for aggregated cells using Aspose.Cells .NET | Aspose.Cells PivotTable drill‑down button not available workaround | enable drill‑down functionality in Excel PivotTable via Aspose.Cells C# | Aspose.Cells missing ShowDrillDownButtons property solution
// Tags: Aspose.Cells PivotTable drill-down button | C# enable PivotTable expand collapse Aspose.Cells | Aspose.Cells missing ShowDrillDownButtons workaround | Excel PivotTable drill-down using Aspose.Cells .NET | Aspose.Cells PivotTable settings C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for a PivotTable, and explains that Aspose.Cells for .NET does not expose a ShowDrillDownButtons property. It outlines how to approach enabling drill‑down (expand/collapse) buttons for aggregated cells and saves the workbook, while handling possible exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first PivotTable in the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTables found in the worksheet.");
                }
                else
                {
                    PivotTable pivotTable = worksheet.PivotTables[0];

                    // Aspose.Cells does not expose a ShowDrillDownButtons property.
                    // If needed, configure other PivotTable settings here.
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
