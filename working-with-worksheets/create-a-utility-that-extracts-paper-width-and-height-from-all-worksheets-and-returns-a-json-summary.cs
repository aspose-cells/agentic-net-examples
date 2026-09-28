// Title: Extract worksheet paper width and height from an Excel workbook and generate a JSON summary with Aspose.Cells for .NET
// AI Prompts: Write a C# method that loads an Excel file with Aspose.Cells, iterates over all worksheets, reads PageSetup.PaperWidth and PageSetup.PaperHeight, and returns an indented JSON string containing the worksheet name and its dimensions. | Extend the method to also capture each worksheet's orientation (Portrait or Landscape) and include it in the JSON output. | Add robust error handling that logs worksheets where the paper size values are default or missing, then continues processing the remaining sheets.
// Common Searches: Aspose.Cells C# get paper width and height for every sheet in a workbook | How to serialize Excel worksheet page setup dimensions to JSON using Aspose.Cells | C# extract page setup paper size from multiple worksheets with Aspose.Cells | Generate JSON report of worksheet paper dimensions with Aspose.Cells .NET
// Tags: aspocells extract worksheet paper size | c# json serialization of pagesetup dimensions | aspocells pagesetup paperwidth paperheight | excel workbook worksheet dimensions extraction | c# generate json summary of worksheet page setup

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Model representing paper size information for a worksheet
    // Loads an Excel workbook with Aspose.Cells, reads each worksheet's PageSetup PaperWidth and PaperHeight (in points), and returns a formatted JSON array that lists the worksheet name together with its width and height values.
    public class WorksheetPaperInfo
    {
        public string WorksheetName { get; set; }
        public double PaperWidth { get; set; }   // Width in points
        public double PaperHeight { get; set; }  // Height in points
    }

    public static class PaperSizeExtractor
    {
        /// <param name="filePath">Full path to the Excel workbook.</param>
        /// <returns>JSON string containing an array of worksheet paper size data.</returns>
        public static string GetPaperSizeSummary(string filePath)
        {
            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The file '{filePath}' was not found.", filePath);
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(filePath);

                // Prepare a list to hold paper size data for each worksheet
                List<WorksheetPaperInfo> summary = new List<WorksheetPaperInfo>();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the PageSetup object for the current worksheet
                    PageSetup pageSetup = sheet.PageSetup;

                    // Collect width and height (in points). If not set, Aspose returns default values.
                    double width = pageSetup.PaperWidth;
                    double height = pageSetup.PaperHeight;

                    // Add the information to the summary list
                    summary.Add(new WorksheetPaperInfo
                    {
                        WorksheetName = sheet.Name,
                        PaperWidth = width,
                        PaperHeight = height
                    });
                }

                // Serialize the summary list to JSON
                JsonSerializerOptions options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                return JsonSerializer.Serialize(summary, options);
            }
            catch (Exception ex)
            {
                // Wrap any exception in a more descriptive message
                throw new InvalidOperationException("Failed to extract paper size information.", ex);
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            // Replace with the path to your Excel file
            string excelPath = @"C:\Path\To\YourWorkbook.xlsx";

            try
            {
                // Get JSON summary of paper sizes
                string jsonSummary = PaperSizeExtractor.GetPaperSizeSummary(excelPath);

                // Output the JSON to console (or handle as needed)
                Console.WriteLine(jsonSummary);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
