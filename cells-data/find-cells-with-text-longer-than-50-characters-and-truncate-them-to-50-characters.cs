// Title: Truncate Excel cell strings longer than 50 characters with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that scans every used cell in all worksheets and replaces any string longer than 50 characters with its first 50 characters. | Create a method that receives a Workbook object and enforces a 50‑character maximum on all text cells, then saves the file. | Generate a loop that iterates through the used range of each worksheet and trims cell values exceeding 50 characters, preserving other data types.
// Common Searches: Aspose.Cells C# truncate cell text to 50 characters | limit Excel cell string length using Aspose.Cells .NET | how to trim long strings in all worksheets with Aspose.Cells | C# code to cut cell values over 50 characters in an Excel file | process used range and shorten cell strings Aspose.Cells
// Tags: cell text truncation Aspose.Cells | string length limit Excel .NET | iterate worksheets used range C# | cell value trimming Aspose.Cells | text length restriction workbook

using System;
using Aspose.Cells;

// The program loads an Excel workbook with Aspose.Cells, iterates through every used cell across all worksheets, truncates any string longer than 50 characters to the first 50 characters, and saves the modified file.
class TruncateLongText
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";
        // Path to save the modified workbook
        string outputPath = "output.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue;

                        // Truncate if longer than 50 characters
                        if (!string.IsNullOrEmpty(text) && text.Length > 50)
                        {
                            cell.PutValue(text.Substring(0, 50));
                        }
                    }
                }
            }
        }

        // Save the modified workbook (save rule)
        workbook.Save(outputPath);
    }
}
