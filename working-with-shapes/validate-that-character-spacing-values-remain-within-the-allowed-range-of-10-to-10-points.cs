// Title: Check and Adjust Cell Font Size (Character Spacing) to Stay Within -10 to 10 Points Using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an Excel workbook, iterates every used cell, reads the Font.Size property, and resets any value outside the -10 to 10 point range. | Create a method that validates character spacing by clamping Font.Size for each cell across all worksheets and then saves the corrected workbook.
// Common Searches: Aspose.Cells how to enforce font size limits in an Excel file C# | C# iterate all cells and correct out‑of‑range character spacing with Aspose.Cells | Validate and clamp cell style properties using Aspose.Cells .NET | Check font size range in each cell of a workbook with Aspose.Cells
// Tags: Aspose.Cells validate cell font size range | C# clamp character spacing Aspose.Cells | Excel workbook style property correction Aspose.Cells | iterate all cells Aspose.Cells .NET | font size bounds enforcement Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, walks through every used cell in all worksheets, reads the Font.Size (used as a proxy for character spacing), clamps any values outside the -10 to 10 point range, writes the corrected style back to the cell, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Allowed font size range in points (used here as a stand‑in for character spacing)
        const double MinSize = -10.0;
        const double MaxSize = 10.0;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the used range of cells to limit iteration
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

            int startRow = usedRange.FirstRow;
            int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
            int startCol = usedRange.FirstColumn;
            int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Retrieve the cell's style to access font properties
                    Style style = cell.GetStyle();

                    // Current font size value (used as a proxy for character spacing)
                    double size = style.Font.Size;

                    // Validate the size value
                    if (size < MinSize || size > MaxSize)
                    {
                        // Clamp to the nearest allowed value
                        double corrected = Math.Max(MinSize, Math.Min(MaxSize, size));

                        // Apply corrected size (cast to int if Font.Size expects int)
                        style.Font.Size = (int)corrected;

                        // Apply the corrected style back to the cell
                        cell.SetStyle(style);
                    }
                }
            }
        }

        try
        {
            // Save the workbook after validation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }
}
