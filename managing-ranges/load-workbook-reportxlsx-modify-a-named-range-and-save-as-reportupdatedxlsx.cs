// Title: Update the first cell of a named range in an Excel workbook and save as a new file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change value of the first cell in a named range and save as new workbook | How to modify a named range cell in an existing Excel file with Aspose.Cells .NET | Aspose.Cells .NET retrieve named range coordinates and update cell value | Save modified Excel workbook under a different name using Aspose.Cells C# | Check if a named range exists before updating its cell with Aspose.Cells
// Tags: named range cell update Aspose.Cells | named range coordinate extraction .NET | workbook save with new filename Aspose.Cells | file existence validation C# Aspose | first cell modification in Excel named range

using System;
using System.IO;
using Aspose.Cells;

// The C# program checks that 'Report.xlsx' exists, loads it with Aspose.Cells, obtains the coordinates of the named range 'MyNamedRange', updates the first cell of that range to "Updated", and saves the result as 'ReportUpdated.xlsx', handling missing files, absent worksheets, and other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "Report.xlsx";
        const string outputPath = "ReportUpdated.xlsx";
        const string rangeName = "MyNamedRange";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range from the workbook's name collection
            Name namedRange = workbook.Worksheets.Names[rangeName];

            if (namedRange != null)
            {
                // Obtain the first cell coordinates of the named range
                int firstRow = namedRange.GetRange().FirstRow;
                int firstColumn = namedRange.GetRange().FirstColumn;

                // Ensure the worksheet exists
                if (workbook.Worksheets.Count > 0)
                {
                    // Access the first cell in the first worksheet (named ranges are workbook‑wide)
                    Cell firstCell = workbook.Worksheets[0].Cells[firstRow, firstColumn];

                    // Modify the cell value
                    firstCell.PutValue("Updated");
                    Console.WriteLine($"Named range '{rangeName}' first cell updated.");
                }
                else
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                }
            }
            else
            {
                Console.WriteLine($"Named range '{rangeName}' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
