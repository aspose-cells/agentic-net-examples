// Title: Replace an Excel workbook’s theme font with a custom font and apply it to every cell using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change workbook theme font to a custom font | set default font for all cells in an existing XLSX file with Aspose.Cells | apply a specific font to every used cell across all worksheets using Aspose.Cells .NET | programmatically replace Excel theme font in C# with Aspose.Cells | update cell styles to use a custom font throughout a workbook via Aspose.Cells
// Tags: replace workbook theme font Aspose.Cells | set default style font C# | apply custom font to all cells Aspose.Cells | iterate worksheets and cells Aspose.Cells | save modified workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX workbook, sets the workbook's default style font to a specified custom font, then iterates through each worksheet and every used cell to assign the same custom font to each cell's style, and finally saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string customFont = "MyCustomFont";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the default font for the workbook (acts as a fallback for cells without explicit font)
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.Font.Name = customFont;
            workbook.DefaultStyle = defaultStyle;

            // Iterate through all worksheets and cells to apply the custom font
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range (max row/column with data)
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        Style style = cell.GetStyle();

                        // Apply the custom font to the cell style
                        style.Font.Name = customFont;

                        // Apply the updated style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
