// Title: How to trim leading and trailing spaces from string cells in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# routine that loads an .xlsx workbook with Aspose.Cells, walks through every used cell, removes surrounding spaces from string values, and writes the cleaned file. | Create a helper method that takes source and destination paths, strips leading and trailing blanks from all text cells in the workbook, and returns the saved file location. | Adjust the sample loop so it only writes back cells whose trimmed text differs from the original, using Aspose.Cells APIs.
// Common Searches: Aspose.Cells C# remove extra spaces from all string cells in an Excel file | C# code to strip whitespace in Excel worksheet using Aspose.Cells library | How to clean up leading and trailing spaces in Excel cells programmatically with Aspose.Cells | Iterate used range and clean cell text in .NET Aspose.Cells example
// Tags: trim whitespace Aspose.Cells C# | remove leading and trailing spaces Excel .NET | iterate used range cells Aspose.Cells | string cell cleanup Aspose.Cells | save modified workbook Aspose.Cells

using System;
using Aspose.Cells;

namespace TrimCellWhitespace
{
    // The program loads an Excel workbook with Aspose.Cells, iterates over the used range, trims leading and trailing spaces from each string cell, updates only cells whose values changed, and saves the cleaned workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (or iterate through all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate over the used range of cells
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        string trimmed = original.Trim();

                        // Update the cell only if trimming changed the value
                        if (!original.Equals(trimmed))
                        {
                            cell.PutValue(trimmed);
                        }
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}
