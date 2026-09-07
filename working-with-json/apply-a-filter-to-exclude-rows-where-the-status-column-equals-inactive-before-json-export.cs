// Title: How to hide rows with "Inactive" status in an Excel worksheet before exporting to JSON using Aspose.Cells for .NET
// AI Prompts: Hide rows where the Status column equals "Inactive" and then save the worksheet as JSON with Aspose.Cells in C#. | Filter out rows with a specific status value before performing a JSON export using Aspose.Cells. | Programmatically exclude rows based on column content when converting an Excel file to JSON with Aspose.Cells. | Use Aspose.Cells to hide rows with a given status and generate a JSON file that omits those rows.
// Common Searches: Aspose.Cells hide rows with specific cell value before JSON export C# | Exclude rows where Status = Inactive when converting Excel to JSON using Aspose.Cells | C# filter Excel data by column value prior to JSON serialization with Aspose.Cells | How to prevent hidden rows from appearing in JSON output from Aspose.Cells | Remove inactive records from Excel to JSON conversion in .NET
// Tags: hide rows based on column value Aspose.Cells | filter Excel rows before JSON export C# | exclude inactive status Aspose.Cells JSON conversion | row visibility control Aspose.Cells | Excel to JSON conversion with row filtering .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, hides any rows where the Status column (index 2) contains "Inactive", and then saves the worksheet as a JSON file. Aspose.Cells automatically omits hidden rows from the exported JSON, providing a filtered output.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.json";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Zero‑based index of the "Status" column (e.g., column C => 2)
        int statusColumnIndex = 2;

        // Determine the last row that contains data
        int lastDataRow = sheet.Cells.MaxDataRow;

        // Loop through data rows (skip header row at index 0)
        for (int row = 1; row <= lastDataRow; row++)
        {
            // Get the cell value in the Status column
            Cell statusCell = sheet.Cells[row, statusColumnIndex];

            // If the status equals "Inactive", hide the entire row
            if (statusCell.StringValue.Equals("Inactive", StringComparison.OrdinalIgnoreCase))
            {
                // Hide the row using Cells.Rows collection
                sheet.Cells.Rows[row].IsHidden = true;
            }
        }

        try
        {
            // Export the worksheet to JSON (hidden rows are excluded)
            workbook.Save(outputPath, SaveFormat.Json);
            Console.WriteLine($"Workbook saved as JSON to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save JSON: {ex.Message}");
        }
    }
}
