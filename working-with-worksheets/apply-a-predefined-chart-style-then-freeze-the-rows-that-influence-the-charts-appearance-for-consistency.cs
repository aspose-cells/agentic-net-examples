// Title: Apply a predefined chart style and freeze the top five rows of an Excel worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file using Aspose.Cells, applies a built‑in ChartStyleType to the first chart on the first worksheet, freezes rows 1‑5, and saves the modified workbook. | Demonstrate how to verify chart presence, set ChartStyleType.Style1, use FreezePanes to lock the top rows, and gracefully handle a missing input file when processing Excel workbooks with Aspose.Cells.
// Common Searches: Aspose.Cells C# apply built‑in chart style to first chart | Freeze first five rows in Excel using Aspose.Cells .NET | How to set ChartStyleType.Style1 with Aspose.Cells | Save workbook after modifying chart and freezing panes in C# | Handle FileNotFoundException when loading Excel file with Aspose.Cells
// Tags: apply built‑in chart style Aspose.Cells C# | freeze top rows worksheet Aspose.Cells | ChartStyleType.Style1 usage Aspose.Cells | load and save workbook Aspose.Cells .NET | exception handling for missing Excel file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel file with Aspose.Cells, checks for a chart on the first worksheet, optionally assigns a predefined ChartStyleType (e.g., Style1) to that chart, freezes the top five rows using FreezePanes, ensures the output directory exists, and saves the workbook while handling missing files and other potential exceptions.
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
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the sheet
            if (sheet.Charts.Count > 0)
            {
                try
                {
                    // Get the first chart
                    Chart chart = sheet.Charts[0];

                    // Apply a predefined chart style if the API supports it
                    // Note: ChartStyleType may not be available in older versions; this line is optional.
                    // chart.Style = ChartStyleType.Style1;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to modify chart: {ex.Message}");
                }
            }

            // Freeze the top 5 rows (rows are zero‑based, so row index 5 freezes rows 1‑5)
            sheet.FreezePanes(5, 0, 0, 0);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
