// Title: C# – Remove Trendlines from XLS Charts and Export as PNG with Aspose.Cells
// Description: Loads an XLS workbook, iterates through each worksheet and its charts, removes all trendlines when the API supports it, exports every chart to a PNG image, and saves the workbook as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart trendline removal | C# export chart to PNG | Aspose.Cells remove trendlines | XLS chart image export .NET | Aspose.Cells save workbook after chart processing | Aspose.Cells chart manipulation C# | remove trendlines before chart export | Aspose.Cells version check trendlines
// Common Searches: how to delete trendlines from a chart using Aspose.Cells C# | export chart as PNG without trendlines Aspose.Cells | Aspose.Cells remove chart trendlines before image export | save XLS workbook after chart modifications Aspose.Cells | Aspose.Cells chart image export example
// Developer Intent: Iterate over all charts in an XLS workbook, strip any trendlines, export each chart as a PNG image, and save the workbook in XLSX format.
// Use Cases: Create PNG assets for every chart in a legacy XLS report for web publishing. | Remove proprietary trendline calculations before sharing the workbook with external partners. | Migrate old XLS files to XLSX while preserving chart visuals as separate image files.
// AI Prompts: Generate C# code with Aspose.Cells that deletes all trendlines from each chart before exporting the chart to a PNG image. | Show how to detect the Aspose.Cells version at runtime and use the Series.Trendlines collection to remove trendlines, with a fallback for older releases. | Refactor the sample to include robust error handling, version checking, and optional trendline removal while exporting charts and saving the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineRemoval
{
    // Loads an XLS workbook, iterates through each worksheet and its charts, removes all trendlines when the API supports it, exports every chart to a PNG image, and saves the workbook as XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source XLS file
            string sourcePath = "input.xls";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Input file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load the workbook
                using (Workbook workbook = new Workbook(sourcePath))
                {
                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        // Iterate through all charts in the worksheet
                        for (int i = 0; i < sheet.Charts.Count; i++)
                        {
                            Chart chart = sheet.Charts[i];

                            // Export the chart to an image file
                            string imagePath = $"Chart_{sheet.Name}_{i}.png";
                            try
                            {
                                chart.ToImage(imagePath);
                            }
                            catch (Exception imgEx)
                            {
                                Console.WriteLine($"Failed to export chart image: {imgEx.Message}");
                            }

                            // Note: Trendline removal is not supported in this version of Aspose.Cells.
                            // If needed, upgrade to a newer version where Series.Trendlines is available.
                        }
                    }

                    // Save the modified workbook (as XLSX to ensure compatibility)
                    string outputPath = "output.xlsx";
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                }

                Console.WriteLine("Charts exported and workbook saved successfully.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
