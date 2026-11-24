// Title: C# utility to extract paper width and height from every worksheet and output a JSON summary using Aspose.Cells
// Common Searches: Aspose.Cells get page setup paper width and height for each worksheet in C# | How to export Excel worksheet paper dimensions to JSON with Aspose.Cells .NET | C# code to list paper size of all sheets in an Excel workbook using Aspose.Cells | Retrieve worksheet print paper size (width/height) via Aspose.Cells API | Generate JSON summary of worksheet page setup dimensions in .NET
// Tags: Aspose.Cells extract worksheet paper size | C# serialize worksheet page setup to JSON | Aspose.Cells iterate all worksheets paper dimensions | PageSetup PaperWidth PaperHeight .NET | Excel workbook paper size extraction using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // DTO for JSON serialization
    // The PaperSizeExtractor class checks the Excel file path, loads the workbook with Aspose.Cells, iterates through each worksheet, reads the PageSetup.PaperWidth and PaperHeight values, stores them in a WorksheetPaperInfo DTO, and returns a formatted JSON array; errors such as missing files are returned as JSON error objects.
    public class WorksheetPaperInfo
    {
        public string WorksheetName { get; set; }
        public double PaperWidth { get; set; }   // Width in points (1 point = 1/72 inch)
        public double PaperHeight { get; set; }  // Height in points
    }

    public static class PaperSizeExtractor
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>JSON string containing an array of worksheet paper size information.</returns>
        public static string ExtractPaperSizes(string filePath)
        {
            // Validate file existence to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                return JsonSerializer.Serialize(
                    new { error = $"File not found: {filePath}" },
                    new JsonSerializerOptions { WriteIndented = true });
            }

            try
            {
                // Load the workbook (Aspose.Cells automatically detects the format)
                var workbook = new Workbook(filePath);

                // Prepare a list to hold the results
                var result = new List<WorksheetPaperInfo>();

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Access the PageSetup object to get paper dimensions
                    PageSetup pageSetup = sheet.PageSetup;

                    // Create a DTO instance for the current worksheet
                    var info = new WorksheetPaperInfo
                    {
                        WorksheetName = sheet.Name,
                        PaperWidth = pageSetup.PaperWidth,
                        PaperHeight = pageSetup.PaperHeight
                    };

                    result.Add(info);
                }

                // Serialize the list to JSON using System.Text.Json
                return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (Exception ex)
            {
                // Return error information as JSON
                return JsonSerializer.Serialize(
                    new { error = ex.Message },
                    new JsonSerializerOptions { WriteIndented = true });
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            // Replace with the path to your Excel file
            string excelPath = @"C:\Temp\SampleWorkbook.xlsx";

            // Extract paper sizes and get JSON summary
            string jsonSummary = PaperSizeExtractor.ExtractPaperSizes(excelPath);

            // Output the JSON to console (or handle as needed)
            Console.WriteLine(jsonSummary);
        }
    }
}
