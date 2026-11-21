// Title: Print all cell formulas from an Excel workbook with Aspose.Cells for .NET (C# example)
// Common Searches: C# code to list every formula in an .xlsx file using Aspose.Cells | How to enumerate formulas across worksheets with Aspose.Cells API | Aspose.Cells read cell formula property in a .NET console application | Extract and display Excel formulas using the Aspose.Cells Workbook class | Iterate through worksheets and output cell formulas in C# with Aspose.Cells
// Tags: Aspose.Cells enumerate worksheet formulas | C# extract Excel cell formulas | Aspose.Cells read formula property | Console output Excel formulas .NET | Workbook iterate worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, loops through each worksheet, and writes to the console the worksheet name followed by every cell that contains a non‑empty Formula, showing the cell address and its formula; errors are caught and reported.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Iterate through all cells and output those that contain formulas
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Console.WriteLine($"{cell.Name}: {cell.Formula}");
                    }
                }

                Console.WriteLine(); // Blank line between worksheets
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
