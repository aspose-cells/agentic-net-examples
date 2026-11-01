// Title: How to loop through only populated rows and columns in an Excel worksheet using Aspose.Cells MaxDataRow and MaxDataColumn in C#
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, determines the last used row and column via MaxDataRow/MaxDataColumn, and iterates only over cells that contain data. | Show an example of skipping empty cells while reading a worksheet with Aspose.Cells, printing each non‑empty cell's address and value. | Create a C# console program that processes a workbook, uses MaxDataRow and MaxDataColumn to define the data bounds, and saves the file after optional modifications.
// Common Searches: aspnet c# read only filled cells in excel using aspose.cells maxdatarow | c# aspocells iterate over used range without empty rows | determine last populated column in worksheet aspocells | skip blank cells when processing excel file with aspocells .net
// Tags: Aspose.Cells iterate used range | MaxDataRow MaxDataColumn C# | skip empty cells Excel .NET | read populated cells Aspose.Cells | determine worksheet data bounds Aspose

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, uses Worksheet.Cells.MaxDataRow and MaxDataColumn to find the last row and column that contain data, loops through that rectangular range, skips cells with null or empty values, prints each populated cell's address and value, and saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the populated data range
            int maxDataRow = worksheet.Cells.MaxDataRow;       // zero‑based index of last row with data
            int maxDataColumn = worksheet.Cells.MaxDataColumn; // zero‑based index of last column with data

            // Iterate only through the populated range
            for (int row = 0; row <= maxDataRow; row++)
            {
                for (int col = 0; col <= maxDataColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Skip cells that have no value
                    if (cell.Value == null || string.IsNullOrEmpty(cell.StringValue))
                        continue;

                    // Example processing: output cell address and value
                    Console.WriteLine($"Cell[{row},{col}] = {cell.Value}");
                }
            }

            // Save the workbook if any modifications were made (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
