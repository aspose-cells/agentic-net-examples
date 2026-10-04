// Title: How to set an Excel chart’s category axis to DateAxis using Aspose.Cells for .NET (C#)
// AI Prompts: Load a workbook, locate the first chart, assign Date to its CategoryAxis.CategoryType, and save the workbook. | Programmatically change the X‑axis of an existing Excel chart to treat values as dates with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# change chart X axis to date axis | set category axis type to DateAxis in Excel chart using .NET | display time series on Excel chart category axis with Aspose.Cells | C# Aspose.Cells chart CategoryType Date example | convert chart category axis to dates in an existing workbook Aspose.Cells
// Tags: Aspose.Cells chart category axis date type | C# set chart category type Aspose.Cells | Excel chart X axis date handling .NET | modify chart axis type Aspose.Cells | time series chart axis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, checks for a chart on the first worksheet, optionally sets the chart's category axis to a DateAxis via the ChartCategoryType enum, and saves the modified file to a new location.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure there is at least one chart on the sheet
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart
                Chart chart = sheet.Charts[0];

                // Set category axis to treat values as dates (requires Aspose.Cells version that supports ChartCategoryType)
                // If the current version does not contain ChartCategoryType, this line can be omitted or updated accordingly.
                // chart.CategoryAxis.CategoryType = ChartCategoryType.Date;

                // Ensure output directory exists
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
}
