// Title: How to truncate Excel cell strings longer than 50 characters using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, scans every used cell, and replaces any string longer than 50 characters with its first 50 characters. | Create a C# routine using Aspose.Cells to iterate all worksheets, detect string cells exceeding 50 characters, truncate them to 50 characters, and save the updated file.
// Common Searches: Aspose.Cells C# truncate cell text to specific length | limit Excel cell string length to 50 characters with Aspose.Cells | C# iterate through all cells in workbook and shorten long strings using Aspose | how to cut off text in Excel cells after 50 characters programmatically | Aspose.Cells replace long string values in cells with substring
// Tags: truncate cell string Aspose.Cells C# | limit Excel cell text length .NET | iterate worksheets Aspose.Cells truncate | substring cell value Aspose.Cells | process used range Aspose.Cells C#

using System;
using Aspose.Cells;

namespace TruncateLongCellText
{
    // C# program that loads an Excel workbook with Aspose.Cells, walks through every worksheet and used cell, truncates any string longer than 50 characters to exactly 50 characters, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range of the worksheet
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain string values
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;
                            if (text != null && text.Length > 50)
                            {
                                // Truncate to the first 50 characters
                                string truncated = text.Substring(0, 50);
                                cell.PutValue(truncated);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
