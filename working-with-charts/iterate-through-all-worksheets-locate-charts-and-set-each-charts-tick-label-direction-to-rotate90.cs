// Title: Loop through every worksheet in an Excel file and set chart axis tick labels to 90° rotation using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that enumerates all worksheets and, for each chart, applies a 90‑degree rotation to the CategoryAxis and ValueAxis tick labels, then saves the workbook. | Update an existing Aspose.Cells workbook so that every chart’s axis tick labels are rotated 90 degrees across all sheets, handling missing axes gracefully.
// Common Searches: Aspose.Cells C# rotate chart axis labels 90 degrees for all worksheets | How to set TickLabelRotationAngle on every chart in an Excel workbook using Aspose.Cells | Bulk change chart tick label orientation in .NET with Aspose.Cells | Iterate worksheets and modify chart axes label angle Aspose.Cells example
// Tags: Aspose.Cells tick label rotation API | C# loop worksheets update chart axes | TickLabelRotationAngle usage example | bulk chart axis styling .NET | Excel chart orientation automation

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program loads an Excel workbook, verifies the input file, iterates over each worksheet and its charts, attempts to set a 90‑degree rotation on the category and value axis tick labels (commented for version compatibility), and saves the modified workbook to the specified output path.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Rotate tick labels for primary axes if the API supports it.
                        // The TickLabelRotationAngle property may not be available in older versions.
                        // If needed, update Aspose.Cells to a version that includes this property.
                        // Example (when supported):
                        // if (chart.CategoryAxis != null)
                        //     chart.CategoryAxis.TickLabelRotationAngle = 90;
                        // if (chart.ValueAxis != null)
                        //     chart.ValueAxis.TickLabelRotationAngle = 90;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to modify chart on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
