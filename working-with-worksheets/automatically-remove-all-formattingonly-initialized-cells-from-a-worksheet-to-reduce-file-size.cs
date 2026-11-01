// Title: How to remove formatting‑only initialized cells from every worksheet in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that scans all worksheets and resets the style of cells that have no value to the workbook's default style. | Create a function that iterates through a workbook, detects empty cells with custom formatting, clears their formatting to reduce file size, and saves the result.
// Common Searches: Aspose.Cells remove custom formatting from blank cells to shrink workbook size | C# iterate all worksheets and clear styles of empty cells using Aspose | How to reset style of cells without values in an Excel file with Aspose.Cells | Reduce Excel file size by deleting formatting‑only cells in .NET
// Tags: clear formatting of empty cells Aspose.Cells | reset cell style to default .NET | optimize workbook size by removing styles | iterate worksheets to clean up formatting Aspose | remove formatting‑only initialized cells C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, loops through each worksheet and every cell, and for cells that contain no value but have a non‑default style, it resets the style to the default before saving the modified workbook, thereby reducing file size.
class RemoveFormattingOnlyCells
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Loop through each cell in the collection
                foreach (Cell cell in cells)
                {
                    // Determine if the cell has no value and a non‑default style (formatting only)
                    if (cell.Value == null)
                    {
                        Style cellStyle = cell.GetStyle();
                        Style defaultStyle = workbook.CreateStyle(); // default (no formatting) style

                        // If the cell's style differs from the default, reset it
                        if (!cellStyle.Equals(defaultStyle))
                        {
                            cell.SetStyle(defaultStyle);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
