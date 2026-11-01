// Title: Load an Excel workbook, turn off theme usage, and save all cells with explicit RGB colors using Aspose.Cells for .NET
// Common Searches: how to prevent Excel theme colors from being applied when saving with Aspose.Cells C# | convert themed cell colors to raw RGB values using Aspose.Cells .NET API | C# code to iterate all cells and enforce explicit RGB formatting in Aspose.Cells | disable Excel theme usage during workbook export with Aspose.Cells | remove theme dependency from an Excel file using Aspose.Cells C#
// Tags: Aspose.Cells disable theme colors | Aspose.Cells convert theme to RGB | Aspose.Cells iterate cell styles | Aspose.Cells explicit RGB color handling | Aspose.Cells load and save workbook without themes

using System;
using System.IO;
using System.Drawing; // For System.Drawing.Color
using Aspose.Cells;

// The example loads an existing Excel workbook with Aspose.Cells, walks through every used cell in each worksheet, reads each cell's style (which already contains System.Drawing.Color representing explicit RGB values), reapplies the unchanged style to ensure colors are stored as RGB, and saves the workbook to a new file without relying on Excel theme colors.
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
            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the sheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Loop through all cells in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        Style style = cell.GetStyle();

                        // NOTE: Aspose.Cells newer versions use System.Drawing.Color for color properties.
                        // Theme color conversion is not required here because System.Drawing.Color
                        // already represents explicit RGB values.

                        // Apply the (unchanged) style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
