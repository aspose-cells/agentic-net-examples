// Title: Set a PivotChart legend to the bottom of an XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that opens an existing XLSX workbook, locates the first chart on the first sheet, assigns Legend.Position = LegendPositionType.Bottom, and writes the changes to a new file. | Provide a .NET snippet that checks for a PivotTable, retrieves its chart, updates the legend to appear at the bottom, and includes try‑catch blocks for file and object errors. | Show how to programmatically reposition a chart's legend to the lower edge using Aspose.Cells Chart.Legend API in C#.
// Common Searches: how to move a chart legend to the lower part of an Excel file using Aspose.Cells in C# | Aspose.Cells example for changing legend placement of a PivotChart in a .xlsx workbook | C# code to set Excel chart legend at bottom with Aspose.Cells library | adjusting chart legend position programmatically with Aspose.Cells .NET | saving workbook after modifying chart legend using Aspose.Cells C#
// Tags: Aspose.Cells chart legend bottom placement | C# modify pivot chart legend Aspose.Cells | Aspose.Cells set legend position XLSX | Aspose.Cells .NET chart legend adjustment | Aspose.Cells workbook chart legend update

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// This C# example uses Aspose.Cells to load an existing XLSX workbook, verifies that a PivotTable and a chart exist on the first worksheet, sets the chart's legend position to the bottom via Legend.Position, and saves the modified workbook while handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one PivotTable on the sheet
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first PivotTable (not used directly in this simplified example)
                PivotTable pivotTable = sheet.PivotTables[0];

                // Find the first chart on the sheet
                if (sheet.Charts.Count > 0)
                {
                    Chart chart = sheet.Charts[0];

                    try
                    {
                        // Set the chart legend position to Bottom
                        chart.Legend.Position = LegendPositionType.Bottom;
                        Console.WriteLine("Chart legend position set to Bottom.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to set legend position: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                }
            }
            else
            {
                Console.WriteLine("No PivotTables found on the first worksheet.");
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
