// Title: Trim leading and trailing whitespace from every string cell in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# function that loads an .xlsx file with Aspose.Cells, walks through each worksheet's data range, trims whitespace from string cells, and saves the cleaned workbook to a new file. | Generate reusable C# code that accepts input and output paths, uses Aspose.Cells to eliminate leading and trailing spaces from all string values, and updates a cell only when the trimmed text differs from the original. | Create a C# snippet demonstrating how to identify string cells, apply String.Trim(), and write the trimmed value back using the Cell.PutValue method in Aspose.Cells.
// Common Searches: aspacells c# remove whitespace from excel cells | how to trim spaces in string cells using Aspose.Cells .NET | loop through worksheet data range and clean string values Aspose.Cells C# example | export cleaned Excel file with Aspose.Cells after removing extra spaces | c# program to trim leading and trailing spaces in Excel file using Aspose
// Tags: Aspose.Cells cell value trimming | Excel workbook string sanitization Aspose.Cells | Selective cell update Aspose.Cells | .NET Excel batch cell processing | Aspose.Cells string cleanup

using Aspose.Cells;
using System;

// // Loads an Excel workbook with Aspose.Cells, iterates over every worksheet's used range, trims leading and trailing whitespace from cells containing strings, writes back only when the value changes, and saves the result to a new .xlsx file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Determine the used range (last row and column with data)
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Loop through all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;

                        // Trim leading and trailing whitespace
                        string trimmed = original.Trim();

                        // Update the cell only if trimming changed the value
                        if (!original.Equals(trimmed))
                        {
                            cell.PutValue(trimmed);
                        }
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
