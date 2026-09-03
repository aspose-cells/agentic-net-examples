// Title: Rename Excel worksheets to "Sheet_{sum}" using Aspose.Cells by adding MaxDataRow and MaxDataColumn values in C#
// AI Prompts: Write a C# program with Aspose.Cells that opens a workbook, calculates MaxDataRow + MaxDataColumn for each worksheet, and sets the worksheet name to "Sheet_{sum}". | Generate code that iterates all worksheets in an .xlsx file, computes the sum of the last used row and column indices, and renames each sheet accordingly using Aspose.Cells for .NET. | Create a .NET console application that loads an Excel file, renames each worksheet based on its used‑range dimensions (row + column), and saves the modified workbook to a new file with Aspose.Cells.
// Common Searches: Aspose.Cells how to set worksheet name based on used range size in C# | C# rename Excel sheet to Sheet_ sum of MaxDataRow and MaxDataColumn | calculate MaxDataRow + MaxDataColumn for each worksheet with Aspose.Cells | automate worksheet naming by data dimensions using Aspose.Cells .NET | rename multiple sheets in a workbook programmatically Aspose.Cells
// Tags: rename worksheet based on used range Aspose.Cells | maxdatarow maxdatacolumn sum sheet naming C# | Aspose.Cells worksheet renaming automation | C# Excel sheet name from data dimensions | Aspose.Cells calculate used range indices

using Aspose.Cells;
using System;
using System.IO;

// Loads an input.xlsx workbook, iterates each worksheet, computes sum = MaxDataRow + MaxDataColumn, renames the sheet to "Sheet_{sum}", and saves the result as output.xlsx while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // MaxDataRow and MaxDataColumn are zero‑based indices of the last used cell
                int maxRow = sheet.Cells.MaxDataRow;
                int maxColumn = sheet.Cells.MaxDataColumn;

                // Calculate the sum of the maximum data row and column
                int sum = maxRow + maxColumn;

                // Rename the worksheet using the calculated sum for quick reference
                sheet.Name = $"Sheet_{sum}";
            }

            // Save the updated workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
