using System;
using System.Collections.Generic;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Represents paper size information for a worksheet
    public class WorksheetPaperInfo
    {
        public string WorksheetName { get; set; }
        public double PaperWidthInches { get; set; }
        public double PaperHeightInches { get; set; }
    }

    public static class PaperSizeExtractor
    {
        /// <summary>
        /// Loads the workbook from the specified path, extracts paper width and height
        /// for each worksheet, and returns a JSON string summarizing the information.
        /// </summary>
        /// <param name="filePath">Path to the Excel file.</param>
        /// <returns>JSON string containing worksheet names with their paper dimensions.</returns>
        public static string GetPaperSizes(string filePath)
        {
            // Load the workbook (uses Aspose.Cells' load constructor)
            Workbook workbook = new Workbook(filePath);

            // Prepare a list to hold paper size data for each worksheet
            List<WorksheetPaperInfo> paperInfoList = new List<WorksheetPaperInfo>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Collect the required information
                WorksheetPaperInfo info = new WorksheetPaperInfo
                {
                    WorksheetName = sheet.Name,
                    PaperWidthInches = pageSetup.PaperWidth,
                    PaperHeightInches = pageSetup.PaperHeight
                };

                paperInfoList.Add(info);
            }

            // Serialize the list to JSON (using System.Text.Json)
            JsonSerializerOptions options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            string jsonResult = JsonSerializer.Serialize(paperInfoList, options);

            return jsonResult;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            string excelPath = "input.xlsx"; // replace with your file path
            string jsonSummary = PaperSizeExtractor.GetPaperSizes(excelPath);
            Console.WriteLine(jsonSummary);
        }
    }
}