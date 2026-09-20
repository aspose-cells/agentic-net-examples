// Title: Log a warning in C# using Aspose.Cells when a worksheet’s MaxDataRow is zero but MaxDataColumn is greater than zero
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets and writes a console warning if the sheet’s MaxDataRow equals 0 while MaxDataColumn is greater than 0. | Update an existing Aspose.Cells workbook processing routine to add a check for column‑only data and output a warning message for each affected worksheet.
// Common Searches: Aspose.Cells how to detect worksheets with MaxDataRow = 0 and MaxDataColumn > 0 in C# | C# Aspose.Cells warning for sheets that have columns but no rows | Check MaxDataRow and MaxDataColumn for each worksheet using Aspose.Cells .NET | Log console warning when Excel sheet contains only column headers with Aspose.Cells | Validate worksheet data presence using MaxDataRow and MaxDataColumn in Aspose.Cells C#
// Tags: log warning for maxdatarow zero aspocells c# | detect column‑only data aspocells worksheet | check maxdatarow maxdatacolumn aspocells | worksheet data validation aspocells c# | iterate worksheets aspocells c#

using Aspose.Cells;
using System;

// // Loads an Excel workbook, iterates each worksheet, and writes a console warning when a sheet's MaxDataRow is zero while its MaxDataColumn is greater than zero.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Retrieve the maximum data row and column indices
            int maxDataRow = sheet.Cells.MaxDataRow;       // Zero‑based index of the last row containing data
            int maxDataColumn = sheet.Cells.MaxDataColumn; // Zero‑based index of the last column containing data

            // Check the condition: MaxDataRow is zero while MaxDataColumn is greater than zero
            if (maxDataRow == 0 && maxDataColumn > 0)
            {
                // Log a warning message
                Console.WriteLine($"Warning: Worksheet '{sheet.Name}' has MaxDataRow = 0 but MaxDataColumn = {maxDataColumn}.");
            }
        }

        // Save the workbook if any changes were made (optional)
        // workbook.Save("output.xlsx");
    }
}
