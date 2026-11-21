// Title: C# code to hide zero values in Excel rows beyond the 100th using Aspose.Cells
// Common Searches: Aspose.Cells C# hide zero values only in rows after row 100 | apply custom number format to suppress zeros in Excel rows greater than 100 using .NET | iterate worksheet cells beyond the 100th row and hide zero display with Aspose.Cells | how to prevent zero values from showing in Excel rows 101 onward in C# | set style to hide numeric zero in Aspose.Cells for rows starting at index 100
// Tags: zero suppression style Aspose.Cells C# | custom number format 0;-0;;@ Aspose.Cells | row based cell styling Aspose.Cells | Excel rows >100 zero formatting .NET | iterate used range Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, creates a style with the custom number format "0;-0;;@" to suppress zero display, then iterates over all used cells in rows 101 onward (row index 100 and higher). Numeric cells equal to zero receive the hide‑zero style, and the workbook is saved to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // Create a style that hides zero values using a custom number format
            var hideZeroStyle = workbook.CreateStyle();
            hideZeroStyle.Custom = "0;-0;;@";

            // Determine the used range limits
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate over rows beyond the 100th row (row index 100 = 101st row)
            for (int row = 100; row <= maxRow; row++)
            {
                // Iterate through all used columns in the current row
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = sheet.Cells[row, col];

                    // Check if the cell contains a numeric zero
                    if (cell.Value != null && cell.Value is double d && d == 0.0)
                    {
                        // Apply the style that hides zero values
                        cell.SetStyle(hideZeroStyle);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
