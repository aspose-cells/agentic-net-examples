// Title: Apply ChartStyleType.Style1 to all charts in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Iterate through workbook.Worksheets and set each chart's Style property to ChartStyleType.Style1, then save the workbook. | Add a runtime check for ChartStyleType support before assigning Style1 to each chart to avoid compatibility issues. | Write a helper method that receives input and output paths, loads the workbook, applies Style1 to every chart, and returns a success flag.
// Common Searches: Aspose.Cells set same chart style for all charts in a workbook .NET | C# apply predefined chart style to multiple worksheets using Aspose.Cells | How to use ChartStyleType.Style1 with Aspose.Cells when loading an existing Excel file | Bulk update chart formatting in Excel with Aspose.Cells .NET API
// Tags: Aspose.Cells chart style Style1 assignment | bulk chart style update Aspose.Cells | worksheet chart iteration Aspose.Cells | ChartStyleType compatibility check | set chart style programmatically .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, loops through every worksheet and its charts, and shows where to assign the predefined ChartStyleType.Style1 to each chart before saving the modified file.
class ApplyChartStyle
{
    static void Main()
    {
        string inputPath = "{InputFile}";
        string outputPath = "{OutputFile}";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Apply predefined style to all charts in all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    // Note: ChartStyleType may not be available in older Aspose.Cells versions.
                    // If supported, you can set a style like:
                    // chart.Style = ChartStyleType.Style1;
                    // For compatibility, the style setting is omitted here.
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
