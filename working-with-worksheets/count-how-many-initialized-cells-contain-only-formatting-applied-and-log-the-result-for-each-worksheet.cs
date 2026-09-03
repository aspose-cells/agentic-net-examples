// Title: Count cells that have only formatting (no data) in each worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loops through every worksheet, identifies cells that contain a style but no value, and prints the count for each sheet. | Create a reusable method that determines whether a Cell's style differs from the default and use it to aggregate formatting‑only cells across a workbook.
// Common Searches: Aspose.Cells count empty cells with formatting applied in C# | C# iterate worksheets and find cells that have style but no data using Aspose.Cells | detect cells that contain only formatting in an Excel workbook with Aspose.Cells .NET | how to tally formatted‑only cells per sheet using Aspose.Cells | Aspose.Cells example to log number of styled empty cells in each worksheet
// Tags: count formatted‑only cells Aspose.Cells | detect non‑default cell style .NET | worksheet iteration Aspose.Cells | cell style detection C# | log formatting cell count per worksheet

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The sample loads an Excel file, iterates each worksheet, determines the used range, and counts cells that have no value but possess any formatting (font, fill, alignment, borders, etc.) via a HasFormatting helper method. It logs the count per worksheet and optionally saves the workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int formattedOnlyCount = 0;

                // Determine the range that may contain initialized cells
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan all cells within the determined range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Count cells that have no value but have any formatting applied
                        if (cell != null && cell.Value == null && HasFormatting(cell))
                        {
                            formattedOnlyCount++;
                        }
                    }
                }

                // Log the count for the current worksheet
                Console.WriteLine($"Worksheet \"{sheet.Name}\": {formattedOnlyCount} cells contain only formatting.");
            }

            // Save the workbook (optional, only if modifications are needed)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Determines whether a cell has any formatting different from the default style
    private static bool HasFormatting(Cell cell)
    {
        Style style = cell.GetStyle();

        // Check common formatting properties
        if (style.Font.IsBold ||
            style.Font.IsItalic ||
            style.Font.Color != Color.Empty ||
            style.Font.Size != 0 ||
            style.ForegroundColor != Color.Empty ||
            style.BackgroundColor != Color.Empty ||
            style.HorizontalAlignment != TextAlignmentType.General ||
            style.VerticalAlignment != TextAlignmentType.Bottom ||
            style.IsTextWrapped)
        {
            return true;
        }

        // Check if any border is applied
        foreach (BorderType bt in Enum.GetValues(typeof(BorderType)))
        {
            if (style.Borders[bt].LineStyle != CellBorderType.None)
                return true;
        }

        return false;
    }
}
