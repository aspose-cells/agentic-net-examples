// Title: C# example: iterate only populated cells in an Excel worksheet using Aspose.Cells MaxDataRow and MaxDataColumn
// Common Searches: C# Aspose.Cells loop through only cells that contain data | How to use MaxDataRow and MaxDataColumn to limit iteration in Aspose.Cells | Skip empty rows and columns when reading Excel with Aspose.Cells .NET | Get last populated row and column in an Aspose.Cells worksheet | Iterate over non‑empty cells in an Excel file using Aspose.Cells C# example
// Tags: maxdatarow maxdatacolumn Aspose.Cells | iterate populated cells C# | skip empty cells Aspose.Cells | determine last data row column Aspose.Cells | excel data range traversal Aspose.Cells | cell null value check Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The code loads an Excel workbook with Aspose.Cells, retrieves the first worksheet, uses Cells.MaxDataRow and Cells.MaxDataColumn to find the last row and column that contain data, and then iterates only over that populated range. Empty cells are skipped by checking for a null Value, and each non‑empty cell's address and string value are written to the console before optionally saving the workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row and column that contain data (zero‑based indices)
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate only over the populated range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Skip cells with no value
                    if (cell.Value == null)
                        continue;

                    // Example processing: write cell address and value to console
                    Console.WriteLine($"Row {row + 1}, Column {col + 1}: {cell.StringValue}");
                }
            }

            // Save the workbook if any changes were made (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
