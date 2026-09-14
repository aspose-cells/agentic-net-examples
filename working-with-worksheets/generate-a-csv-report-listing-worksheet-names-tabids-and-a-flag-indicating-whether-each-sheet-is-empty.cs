// Title: Create a CSV report listing each worksheet’s name, TabId, and empty‑sheet status with Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console program that loads an Excel file using Aspose.Cells, iterates all worksheets, and writes a CSV file containing the worksheet name, its TabId, and a true/false flag indicating whether the sheet is empty. | Add a utility method that correctly escapes commas, double quotes, and line breaks in worksheet names when generating CSV rows. | Modify the example to return the CSV data as a MemoryStream or string instead of saving it directly to disk, enabling further processing.
// Common Searches: Aspose.Cells C# export worksheet names and TabId to CSV | How to check if a worksheet is empty using Aspose.Cells .NET | Generate Excel worksheet metadata report with Aspose.Cells and save as CSV | C# code to list all sheet TabIds in an Excel workbook using Aspose.Cells | CSV escaping rules for worksheet names when using Aspose.Cells
// Tags: Aspose.Cells export worksheet metadata to CSV | C# retrieve worksheet TabId Aspose.Cells | detect empty worksheet Aspose.Cells .NET | CSV field escaping for Excel sheet names C# | generate worksheet report Aspose.Cells workbook

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Generates a CSV file that lists each worksheet’s name, its TabId, and a boolean flag indicating if the sheet contains no data (MaxDataRow == -1), with proper CSV escaping for special characters in sheet names.
class WorksheetReportGenerator
{
    static void Main(string[] args)
    {
        // Input Excel file path (change as needed)
        string excelFilePath = "input.xlsx";

        // Output CSV file path
        string csvReportPath = "WorksheetReport.csv";

        // Load the workbook
        Workbook workbook = new Workbook(excelFilePath);

        // Prepare a StringBuilder for CSV content
        StringBuilder csvBuilder = new StringBuilder();

        // Write CSV header
        csvBuilder.AppendLine("WorksheetName,TabId,IsEmpty");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get worksheet name
            string sheetName = sheet.Name;

            // Get TabId (unique identifier for the worksheet tab)
            int tabId = sheet.TabId;

            // Determine if the worksheet is empty.
            // MaxDataRow returns -1 when there are no cells with data.
            bool isEmpty = sheet.Cells.MaxDataRow == -1;

            // Build CSV line (escaping commas if necessary)
            string csvLine = $"{EscapeCsv(sheetName)},{tabId},{isEmpty}";
            csvBuilder.AppendLine(csvLine);
        }

        // Write the CSV content to file
        File.WriteAllText(csvReportPath, csvBuilder.ToString(), Encoding.UTF8);

        Console.WriteLine($"CSV report generated at: {csvReportPath}");
    }

    // Helper method to escape CSV fields that may contain commas or quotes
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            // Escape double quotes by doubling them
            string escaped = field.Replace("\"", "\"\"");
            return $"\"{escaped}\"";
        }
        return field;
    }
}
