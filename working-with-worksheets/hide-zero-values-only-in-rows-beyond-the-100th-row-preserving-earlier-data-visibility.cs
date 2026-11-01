// Title: Hide zero values in Excel rows beyond the 100th row with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to traverse rows 101 onward, detect cells whose numeric value equals 0, and assign the custom format '0;-0;;@' to conceal those zeros. | Programmatically hide zero entries in the lower part of an Excel sheet by applying a style that omits the zero section of a custom number format, then save the workbook.
// Common Searches: asp.net hide zero values in Excel after the 100th row using Aspose.Cells | c# Aspose.Cells hide zeros only in rows greater than 100 | how to suppress zero display in later rows of an Excel worksheet with .NET | custom number format to hide zeros in specific rows using Aspose.Cells
// Tags: Aspose.Cells custom number format zero suppression | C# zero value concealment in Excel rows after 100 | apply cell style for zero suppression Aspose | iterate worksheet rows beyond row 100 Aspose.Cells | Excel zero display control .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing workbook, iterates from row 101 to the last used row, applies the custom number format '0;-0;;@' to any cell containing the numeric value 0 to hide those zeros, and saves the updated file.
class HideZeroValuesBeyondRow100
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of rows and columns
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate over rows beyond the 100th row (row index 100 corresponds to row number 101)
            for (int row = 100; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Check if the cell contains a numeric zero value
                    if (cell.Value != null && cell.Value is double d && d == 0.0)
                    {
                        // Apply a custom number format that hides zero values
                        Style style = cell.GetStyle();
                        // Format: positive;negative;zero;text -> zero part left empty hides zeros
                        style.Custom = "0;-0;;@";
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
