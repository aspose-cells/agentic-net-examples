// Title: Count formatting‑only cells (no data) in each worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loops through every worksheet, examines each cell in the used range, and increments a counter for cells that have a style different from the default style but contain no value, then output the count per worksheet. | Modify an existing Aspose.Cells workbook‑processing script to ignore completely empty cells and report the number of cells that contain only formatting for each sheet, logging the results to the console.
// Common Searches: Aspose.Cells how to count cells that have formatting but no value in each worksheet | C# count formatting‑only cells in Excel using Aspose.Cells library | iterate used range and detect non‑default style without data Aspose.Cells .NET | log number of cells with only style per sheet Aspose.Cells | skip empty cells and count formatted cells in workbook Aspose.Cells C#
// Tags: count formatting‑only cells Aspose.Cells | detect non‑default cell style without value C# | iterate used range worksheet Aspose.Cells | log cell formatting count per worksheet | skip empty cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, determines the used range of each worksheet, and counts cells that have no value but a style different from the workbook's default. It then writes the count of formatting‑only cells for each sheet to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Get the default style for comparison
            var defaultStyle = workbook.DefaultStyle;

            // Iterate through each worksheet
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                int formattedOnlyCellCount = 0;

                // Determine the used range
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxColumn = worksheet.Cells.MaxDataColumn;

                // If the sheet has data
                if (maxRow >= 0 && maxColumn >= 0)
                {
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = worksheet.Cells[row, col];

                            // Skip cells that are completely empty (no value and default style)
                            bool isCompletelyEmpty = cell.Value == null &&
                                                     cell.Type == CellValueType.IsNull &&
                                                     cell.GetStyle().Equals(defaultStyle);
                            if (isCompletelyEmpty)
                                continue;

                            // Determine if the cell has no value
                            bool hasNoValue = cell.Value == null ||
                                              cell.Type == CellValueType.IsNull;

                            // Determine if the cell has formatting different from the default
                            bool hasFormatting = !cell.GetStyle().Equals(defaultStyle);

                            // Count cells that have formatting but no value
                            if (hasNoValue && hasFormatting)
                                formattedOnlyCellCount++;
                        }
                    }
                }

                Console.WriteLine($"Worksheet \"{worksheet.Name}\": {formattedOnlyCellCount} cells contain only formatting.");
            }

            // Optional: save if modifications were made
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
