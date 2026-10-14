// Title: Generate a CSV report of worksheet names, TabIds, and empty‑sheet status with Aspose.Cells in C#
// AI Prompts: Create a C# console program that opens an Excel workbook using Aspose.Cells, loops through all worksheets, and writes a CSV line with the worksheet’s name, its TabId, and a boolean indicating if the sheet contains no data. | Add a helper method that correctly escapes commas, quotes, and line breaks in CSV fields, then save the assembled report to a user‑defined file path.
// Common Searches: aspnet aspose.cells retrieve worksheet TabId and empty check | c# export worksheet list to csv with tab identifiers | how to determine if an Excel sheet is empty using Aspose.Cells | generate worksheet metadata csv file from workbook in .NET | c# Aspose.Cells iterate worksheets and write report
// Tags: Aspose.Cells worksheet TabId extraction | detect empty worksheet Aspose.Cells | worksheet information CSV export C# | CSV special character handling C# | enumerate workbook worksheets Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvReport
{
    // The example loads an Excel workbook with Aspose.Cells, iterates each worksheet to obtain its name and TabId, checks emptiness via MaxDataRow/MaxDataColumn, builds a CSV with proper escaping for special characters, and writes the report to a specified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path (change as needed)
            string excelFilePath = "input.xlsx";

            // Output CSV file path
            string csvReportPath = "WorksheetReport.csv";

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(excelFilePath);

            // Prepare a StringBuilder to build CSV content
            StringBuilder csvBuilder = new StringBuilder();

            // Write CSV header
            csvBuilder.AppendLine("WorksheetName,TabId,IsEmpty");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get worksheet name
                string sheetName = sheet.Name;

                // Get TabId (unique identifier for the worksheet tab)
                int tabId = sheet.TabId;

                // Determine if the worksheet is empty.
                // A worksheet is considered empty when it has no data cells.
                // MaxDataRow and MaxDataColumn return -1 when there are no data cells.
                bool isEmpty = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

                // Append the information as a CSV line
                csvBuilder.AppendLine($"{EscapeCsv(sheetName)},{tabId},{isEmpty}");
            }

            // Write the CSV content to the output file
            File.WriteAllText(csvReportPath, csvBuilder.ToString(), Encoding.UTF8);

            Console.WriteLine($"CSV report generated at: {Path.GetFullPath(csvReportPath)}");
        }

        private static string EscapeCsv(string field)
        {
            if (field.Contains("\""))
                field = field.Replace("\"", "\"\"");

            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
                field = $"\"{field}\"";

            return field;
        }
    }
}
