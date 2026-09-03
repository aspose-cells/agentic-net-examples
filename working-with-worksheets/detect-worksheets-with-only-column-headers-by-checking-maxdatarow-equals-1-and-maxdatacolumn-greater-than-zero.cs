// Title: Identify Excel worksheets that contain only column headers (no data rows) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that iterates a workbook and prints the names of sheets where Cells.MaxDataRow is -1 and Cells.MaxDataColumn > 0. | Show how to use the Cells.MaxDataRow and Cells.MaxDataColumn properties to filter out worksheets that have only header rows in an Excel file. | Provide a snippet that logs each worksheet name that lacks data rows but includes at least one column header using Aspose.Cells for .NET.
// Common Searches: aspocells c# find worksheets that only have header rows | how to detect Excel sheets with no data rows but with column headers using Aspose.Cells | maxdatarow -1 and maxdatacolumn > 0 example in Aspose.Cells | list worksheets containing only column headers in a .xlsx file with Aspose.Cells .NET
// Tags: header-only worksheet detection Aspose.Cells | Cells.MaxDataRow empty data rows check | filter worksheets by MaxDataColumn in .NET | identify sheets with only column headers C# | Excel workbook header-only sheet detection Aspose

using Aspose.Cells;
using System;

// // Loads an Excel workbook, iterates each worksheet, checks Cells.MaxDataRow and Cells.MaxDataColumn, and prints the names of worksheets that have no data rows (MaxDataRow == -1) but contain at least one header column (MaxDataColumn > 0).
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the maximum data row and column indices
            int maxDataRow = sheet.Cells.MaxDataRow;
            int maxDataColumn = sheet.Cells.MaxDataColumn;

            // Detect worksheets that contain only column headers
            // (no data rows, but at least one header column)
            if (maxDataRow == -1 && maxDataColumn > 0)
            {
                Console.WriteLine($"Worksheet \"{sheet.Name}\" has only column headers.");
            }
        }

        // Save the workbook if any modifications are made (optional)
        // workbook.Save("output.xlsx");
    }
}
