// Title: C# utility to extract worksheet paper size (width & height) from an Excel workbook and output JSON with Aspose.Cells
// Description: Loads an Excel file using Aspose.Cells, walks through every worksheet, reads the PageSetup properties that return the page dimensions in inches, stores the sheet name together with its width and height, and serializes the collection into a readable JSON string.
// Keywords: Aspose.Cells worksheet page size | C# extract Excel paper dimensions | JSON export of sheet size | PageSetup PaperWidth PaperHeight | .NET Excel print layout metadata
// Common Searches: how to get page dimensions of each sheet with Aspose.Cells | C# code to list worksheet paper size in inches | export Excel sheet size to JSON using .NET | retrieve paper width height from workbook programmatically
// Developer Intent: Obtain the width and height of every worksheet’s printable area and return the data as a JSON payload.
// Use Cases: Generate a layout report that shows page size before bulk printing | Validate consistency of paper settings across all sheets in a workbook | Feed sheet‑size metadata to downstream services or documentation pipelines
// AI Prompts: Create a C# method that reads PageSetup.PaperWidth and PaperHeight for each worksheet and returns an indented JSON array. | Show how to loop through all worksheets in a workbook, capture their names and page dimensions, and serialize the result with System.Text.Json. | Explain how to extend the JSON output to include orientation, margins, and scaling factors.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Simple DTO to hold paper size information for a worksheet
    // Loads an Excel file using Aspose.Cells, walks through every worksheet, reads the PageSetup properties that return the page dimensions in inches, stores the sheet name together with its width and height, and serializes the collection into a readable JSON string.
    public class WorksheetPaperSize
    {
        public string WorksheetName { get; set; }
        public double PaperWidthInches { get; set; }
        public double PaperHeightInches { get; set; }
    }

    public static class PaperSizeExtractor
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>JSON string containing an array of worksheet paper size information.</returns>
        public static string ExtractPaperSizes(string filePath)
        {
            // Load the workbook (uses Aspose.Cells' built‑in load functionality)
            Workbook workbook = new Workbook(filePath);

            // Prepare a list to hold the results
            List<WorksheetPaperSize> result = new List<WorksheetPaperSize>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Gather width and height (both are read‑only properties returning inches)
                double width = pageSetup.PaperWidth;
                double height = pageSetup.PaperHeight;

                // Add the information to the result list
                result.Add(new WorksheetPaperSize
                {
                    WorksheetName = sheet.Name,
                    PaperWidthInches = width,
                    PaperHeightInches = height
                });
            }

            // Serialize the list to JSON (indented for readability)
            JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(result, options);
            return json;
        }

        // Example entry point demonstrating usage
        public static void Main()
        {
            // Path to the Excel file to analyze
            string inputPath = "sample.xlsx";

            // Extract paper size information and obtain JSON summary
            string jsonSummary = ExtractPaperSizes(inputPath);

            // Output the JSON to console (or write to a file as needed)
            Console.WriteLine(jsonSummary);
        }
    }
}
