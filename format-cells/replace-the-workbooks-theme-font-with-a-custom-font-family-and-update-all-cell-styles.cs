// Title: How to replace an Excel workbook’s theme font with a custom font and apply it to all cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change workbook default font and update existing cell styles | replace theme font in an existing .xlsx file programmatically with Aspose.Cells | set custom font for every cell in a loaded workbook using Aspose.Cells .NET | iterate through worksheets to apply a new font to all cells with Aspose.Cells | C# code to change Excel workbook theme font to Arial Unicode MS using Aspose.Cells
// Tags: Aspose.Cells set workbook default font C# | apply custom font to all cells Aspose.Cells | replace Excel theme font programmatically .NET | iterate worksheets cells change font Aspose.Cells | update cell styles with custom font Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX workbook, sets a custom default font, iterates over every worksheet, row, and column to assign the new font to each cell’s style, and saves the modified file, demonstrating full‑workbook theme font replacement with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        const string customFont = "Arial Unicode MS";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the default font for the workbook (applies to new styles)
            workbook.DefaultStyle.Font.Name = customFont;

            // Update all cell styles to use the custom font
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range limits
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        Style style = cell.GetStyle();

                        // Set the font name to the custom font
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
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
