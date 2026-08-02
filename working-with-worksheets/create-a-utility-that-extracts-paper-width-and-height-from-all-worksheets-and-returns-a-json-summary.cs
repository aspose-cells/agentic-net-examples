// Title: Extract worksheet paper size (width & height) to JSON with Aspose.Cells for .NET
// Description: C# utility that loads an Excel workbook using Aspose.Cells, reads each worksheet's PageSetup.PaperWidth and PaperHeight (in inches), and returns a formatted JSON array containing the sheet name and its dimensions.
// Keywords: Aspose.Cells paper size | C# worksheet page setup | extract Excel paper width height | JSON export Aspose.Cells | PageSetup.PaperWidth | PageSetup.PaperHeight | print layout audit .NET | Excel sheet dimensions JSON
// Common Searches: how to get paper width and height of each Excel sheet using Aspose.Cells | C# extract worksheet page setup dimensions to JSON | Aspose.Cells get page size for all worksheets | convert Excel sheet print size to JSON in .NET | retrieve paper dimensions from workbook with Aspose
// Developer Intent: Read the PaperWidth and PaperHeight of every worksheet in a workbook and output the data as a JSON string.
// Use Cases: Generate a print‑layout audit report that lists each sheet’s paper size before printing. | Provide a JSON configuration for a reporting service that needs exact page dimensions per worksheet. | Quickly display page‑setup settings in the console for developers troubleshooting mismatched print layouts.
// AI Prompts: Write a C# method using Aspose.Cells that collects PaperWidth and PaperHeight from all worksheets and returns an indented JSON string. | Extend the extractor to include PageSetup.Orientation, margins, and header/footer settings in the JSON output. | Create a unit test that validates the JSON produced by GetPaperSizesJson against expected dimensions for a sample workbook.

using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Represents paper dimensions for a worksheet
    // C# utility that loads an Excel workbook using Aspose.Cells, reads each worksheet's PageSetup.PaperWidth and PaperHeight (in inches), and returns a formatted JSON array containing the sheet name and its dimensions.
    public class WorksheetPaperInfo
    {
        public string WorksheetName { get; set; }
        public double WidthInches { get; set; }
        public double HeightInches { get; set; }
    }

    public static class PaperSizeExtractor
    {
        // Extracts paper width and height from all worksheets in the given workbook file
        // and returns a JSON string summarizing the information.
        public static string GetPaperSizesJson(string workbookPath)
        {
            // Load the workbook from the specified file path
            Workbook workbook = new Workbook(workbookPath);

            var paperInfoList = new List<WorksheetPaperInfo>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup of the worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Retrieve paper dimensions (in inches)
                double width = pageSetup.PaperWidth;
                double height = pageSetup.PaperHeight;

                // Store the information
                paperInfoList.Add(new WorksheetPaperInfo
                {
                    WorksheetName = sheet.Name,
                    WidthInches = width,
                    HeightInches = height
                });
            }

            // Serialize the list to a formatted JSON string
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonResult = JsonSerializer.Serialize(paperInfoList, jsonOptions);

            return jsonResult;
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to analyze
            string filePath = "input.xlsx";

            // Get JSON summary of paper sizes
            string jsonSummary = PaperSizeExtractor.GetPaperSizesJson(filePath);

            // Output the JSON to console
            Console.WriteLine(jsonSummary);
        }
    }
}
