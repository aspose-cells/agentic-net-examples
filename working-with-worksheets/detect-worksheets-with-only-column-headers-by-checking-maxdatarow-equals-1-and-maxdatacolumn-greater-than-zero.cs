// Title: Identify Excel worksheets that contain only column headers (no data rows) using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans a workbook and prints the names of sheets where Cells.MaxDataRow is -1 and Cells.MaxDataColumn is greater than zero. | Show how to loop through all worksheets in an Excel file and flag those that have only header columns using the MaxDataRow and MaxDataColumn properties. | Create a .NET console example that reports worksheets that consist solely of column titles without any data rows.
// Common Searches: aspocells find sheets with only headers no data rows | c# check MaxDataRow -1 and MaxDataColumn > 0 in Excel workbook | list worksheets that contain only column titles using Aspose.Cells | detect header‑only worksheets in .NET Excel processing
// Tags: Aspose.Cells header‑only worksheet detection | Cells.MaxDataRow -1 check | Cells.MaxDataColumn header presence | iterate workbook worksheets for header rows | C# identify sheets without data rows

using System;
using Aspose.Cells;

// The program loads an Excel workbook, iterates each worksheet, evaluates Cells.MaxDataRow and Cells.MaxDataColumn, and outputs the names of worksheets that have no data rows (MaxDataRow == -1) but contain at least one column header (MaxDataColumn > 0).
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the maximum data row and column indices
            int maxDataRow = sheet.Cells.MaxDataRow;       // Returns -1 if no data rows exist
            int maxDataColumn = sheet.Cells.MaxDataColumn; // Returns -1 if no data columns exist

            // Detect worksheets that contain only column headers
            // Condition: no data rows (MaxDataRow == -1) and at least one header column (MaxDataColumn > 0)
            if (maxDataRow == -1 && maxDataColumn > 0)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" contains only column headers.");
            }
        }

        // Optional: save the workbook if modifications were made
        // workbook.Save("output.xlsx");
    }
}
