// Title: Reformat MM/dd/yyyy date strings to ISO 8601 (yyyy-MM-dd) across all worksheets in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Scan every string cell in a workbook and replace dates matching MM/dd/yyyy with yyyy-MM-dd using Aspose.Cells C#. | Apply a regular expression to convert US‑style dates to ISO 8601 format while iterating through all worksheets in an Excel file. | Update an Excel workbook by bulk‑editing cell values with a regex date transformation and save the result with Aspose.Cells.
// Common Searches: C# Aspose.Cells replace US date format with ISO 8601 in all sheets | how to use regex to change mm/dd/yyyy to yyyy-mm-dd in an Excel workbook with Aspose | iterate through cells in Aspose.Cells and reformat date strings | bulk update string cells containing dates in .xlsx using Aspose.Cells and C# | convert date strings to ISO 8601 while saving workbook with Aspose.Cells
// Tags: regex date conversion Aspose.Cells C# | bulk cell string replacement Excel .xlsx | reformat US date strings to ISO 8601 Aspose | iterate worksheets update cell values Aspose.Cells | save modified workbook Aspose.Cells .NET

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through every worksheet and cell, identifies string cells that match the MM/dd/yyyy pattern, rewrites each matched date to ISO 8601 (yyyy‑MM‑dd) using a regular expression, updates the cell only when a change occurs, and saves the transformed workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define the regex pattern for dates (e.g., MM/dd/yyyy)
        Regex datePattern = new Regex(@"\b(\d{2})/(\d{2})/(\d{4})\b");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string values (skip empty or non‑string cells)
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;

                        // Replace matched dates with ISO 8601 format (yyyy-MM-dd)
                        string replaced = datePattern.Replace(original, m =>
                        {
                            // m.Groups: 1 = month, 2 = day, 3 = year
                            return $"{m.Groups[3].Value}-{m.Groups[1].Value}-{m.Groups[2].Value}";
                        });

                        // Update the cell only if a replacement occurred
                        if (!original.Equals(replaced))
                        {
                            cell.PutValue(replaced);
                        }
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
