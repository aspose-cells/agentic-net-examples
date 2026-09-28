// Title: Replace "Draft" cells with "Final" in a specific named range using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, locates the named range "PublishRange", and changes every cell whose exact value is "Draft" to "Final", then saves the file. | Write a reusable method that iterates over all cells in a given Aspose.Cells named range and replaces a target string with a new string, handling only string‑type cells.
// Common Searches: Aspose.Cells C# replace specific text in a named range | How to change 'Draft' to 'Final' in an Excel named range using .NET | Iterate through cells of a named range with Aspose.Cells | Update cell values in a publishing range programmatically in C# | Find and replace exact string in Excel named range Aspose.Cells
// Tags: replace string in named range Aspose.Cells | iterate cells in Excel named range C# | Aspose.Cells conditional cell update | load workbook and save after modification .NET | publishing range cell replacement Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, retrieves the named range "PublishRange", replaces any cell containing exactly "Draft" with "Final", and saves the result as output.xlsx using Aspose.Cells for .NET.
class ReplaceDraftWithFinal
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the range used for publishing
            string rangeName = "PublishRange";

            // Retrieve the named range (fully qualified to avoid ambiguity)
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(rangeName);
            if (range == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Iterate through all cells in the range
            for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
            {
                for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
                {
                    // Access the cell via the worksheet that owns the range
                    Cell cell = range.Worksheet.Cells[row, col];

                    // Replace exact "Draft" with "Final"
                    if (cell.Type == CellValueType.IsString && cell.StringValue == "Draft")
                    {
                        cell.PutValue("Final");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
