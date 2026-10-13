// Title: How to split comma‑separated strings in Excel column B into separate columns using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, reads each cell in column B, splits the string on commas, trims each part, writes the parts into consecutive cells starting at column C, and saves the updated file. | Create a .NET method that uses Aspose.Cells to iterate over column B, parse comma‑delimited text, populate adjacent columns with the trimmed elements, and preserve existing worksheet data.
// Common Searches: Aspose.Cells C# split values in a single column by comma and expand to multiple columns | programmatically separate comma delimited list in Excel column using Aspose.Cells .NET | C# example to parse comma separated strings in Excel and write each part to next columns | how to use Aspose.Cells to convert a CSV list inside a cell into separate cells
// Tags: Aspose.Cells split cell by delimiter | C# write split values to adjacent columns | Excel column B to multiple columns Aspose | process comma delimited strings Aspose.Cells | dynamic cell insertion based on split parts .NET

using Aspose.Cells;
using System;

// Loads an Excel workbook, iterates through each row of column B, splits non‑empty string values on commas, trims each segment, writes the segments into consecutive cells beginning with column C, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Find the last row that contains data in the sheet
        int lastRow = cells.MaxDataRow;

        // Iterate through each row in column B (index 1)
        for (int row = 0; row <= lastRow; row++)
        {
            Cell cell = cells[row, 1]; // Column B

            // Process only string cells that are not empty
            if (cell.Type == CellValueType.IsString && !string.IsNullOrWhiteSpace(cell.StringValue))
            {
                // Split the cell value by comma delimiter
                string[] parts = cell.StringValue.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                // Trim each part and write it to subsequent columns starting from column C (index 2)
                for (int i = 0; i < parts.Length; i++)
                {
                    string trimmed = parts[i].Trim();
                    cells[row, 2 + i].PutValue(trimmed);
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
