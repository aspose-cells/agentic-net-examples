// Title: Batch updating linked picture and chart shapes across multiple worksheets using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through every worksheet in a workbook and assigns a new LinkedCell range to each picture shape based on a dictionary of sheet names to ranges. | Create a routine that updates the Values property of all series in chart shapes to the corresponding worksheet range from the same mapping. | Add try‑catch blocks that log the name of any shape that fails to update while allowing the rest of the workbook to be processed.
// Common Searches: Aspose.Cells C# change linked cell of picture shapes in all sheets | How to set chart series source range for multiple worksheets with Aspose.Cells | Batch modify linked pictures and charts in an Excel file using Aspose.Cells .NET | Dictionary based worksheet to range mapping for shape updates in Aspose.Cells
// Tags: batch update linked picture shapes Aspose.Cells | set chart series values range Aspose.Cells | iterate worksheets modify shape data source .NET | worksheet name to range dictionary Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Charts;

// Loads a workbook, uses a dictionary to map worksheet names to new range strings, iterates each worksheet's shapes, updates picture LinkedCell and chart series Values to the mapped range, logs any failures, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook or create a new one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                Console.WriteLine($"Input file '{inputPath}' not found. A new workbook has been created.");
            }

            // Mapping of worksheet names to their new data source ranges.
            var newDataSources = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "Sheet1", "Sheet1!$A$1:$B$10" },
                { "Sheet2", "Sheet2!$C$1:$D$15" }
                // Add more mappings as needed.
            };

            // Iterate through worksheets and update linked pictures and charts.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (!newDataSources.TryGetValue(sheet.Name, out string newRange))
                    continue; // No mapping for this sheet.

                foreach (Shape shape in sheet.Shapes)
                {
                    // Update linked picture shapes.
                    if (shape is Picture picture)
                    {
                        try
                        {
                            picture.LinkedCell = newRange;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to set linked cell for picture '{picture.Name}': {ex.Message}");
                        }
                    }

                    // Update chart shapes.
                    if (shape is ChartShape chartShape)
                    {
                        Chart chart = chartShape.Chart;
                        foreach (Series series in chart.NSeries)
                        {
                            series.Values = newRange;
                        }
                    }
                }
            }

            // Save the modified workbook.
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
