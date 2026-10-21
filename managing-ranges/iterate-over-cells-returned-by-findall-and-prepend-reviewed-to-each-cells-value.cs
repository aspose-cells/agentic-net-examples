// Title: Prepend "Reviewed:" to every non‑empty cell in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells prepend text to each cell in used range | How to add a prefix to all non‑empty Excel cells with Aspose.Cells .NET | Iterate through worksheet cells and modify values using Aspose.Cells C# | Bulk update cell values in an .xlsx file with Aspose.Cells | Aspose.Cells C# prepend string to cell values while skipping blanks
// Tags: prepend text to Excel cells Aspose.Cells | bulk modify cell values .xlsx C# | iterate used range Aspose.Cells | skip empty cells Aspose.Cells C# | update cell value with custom prefix .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, iterates over the used range of the first worksheet, skips empty cells, prepends the string "Reviewed:" to each cell's existing value, and saves the result to a new Excel file.
class Program
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxColumn = sheet.Cells.MaxDataColumn;

            // Iterate through each used cell and prepend "Reviewed:" to its current value
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Skip empty cells
                    if (cell.Value == null)
                        continue;

                    // Get the existing string representation of the cell's value
                    string originalValue = cell.StringValue;

                    // Prepend the prefix and write back to the cell
                    cell.PutValue("Reviewed:" + originalValue);
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
