// Title: C# – Create CSV summary of worksheet names, Tab IDs, and empty‑sheet status with Aspose.Cells
// Description: This C# example loads an Excel file using Aspose.Cells, walks through every worksheet, captures the sheet’s name and TabId, evaluates emptiness by checking MaxDataRow and MaxDataColumn, and writes the results to a CSV file with proper escaping. The workbook can be saved unchanged after the report is generated.
// Keywords: Aspose.Cells | .NET | C# | CSV report | worksheet name | TabId | empty worksheet detection | MaxDataRow | MaxDataColumn | Excel workbook analysis | export worksheet list
// Common Searches: Aspose.Cells list worksheets CSV | C# get TabId of Excel sheets | detect empty worksheets Aspose.Cells | export workbook metadata to CSV | generate worksheet summary .NET
// Developer Intent: Produce a CSV file that enumerates each worksheet’s name, its TabId, and a boolean indicating whether the sheet contains any data.
// Use Cases: Quick audit of workbook structure before data processing | Provide non‑technical stakeholders with a concise sheet inventory | Skip blank worksheets during bulk import operations | Log worksheet metadata for automated quality checks
// AI Prompts: Write a function that returns a DataTable with columns WorksheetName, TabId, IsEmpty using Aspose.Cells. | Extend the sample to also record the total data rows and columns for each sheet in the CSV. | Add comprehensive error handling for missing files, permission issues, and log each step of the CSV generation.

using System;
using System.IO;
using Aspose.Cells;

// This C# example loads an Excel file using Aspose.Cells, walks through every worksheet, captures the sheet’s name and TabId, evaluates emptiness by checking MaxDataRow and MaxDataColumn, and writes the results to a CSV file with proper escaping. The workbook can be saved unchanged after the report is generated.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile); // load rule

        // Path for the CSV report
        string csvFile = "WorksheetReport.csv";

        // Create CSV and write header
        using (StreamWriter writer = new StreamWriter(csvFile))
        {
            writer.WriteLine("WorksheetName,TabId,IsEmpty");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                string name = sheet.Name;
                int tabId = sheet.TabId;

                // Determine if the sheet is empty.
                // A sheet is considered empty when it has no data rows and no data columns.
                bool isEmpty = sheet.Cells.MaxDataRow == 0 && sheet.Cells.MaxDataColumn == 0;

                // Write CSV line (escape name if needed)
                writer.WriteLine($"{EscapeCsv(name)},{tabId},{isEmpty}");
            }
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx"); // save rule
    }

    // Helper to escape CSV fields containing commas, quotes, or newlines
    static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
