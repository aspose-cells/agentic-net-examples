// Title: Validate that Excel cell comments (tooltips) are shown only when the cell text is truncated using Aspose.Cells for .NET
// Common Searches: C# code to check if a cell's comment should appear only when its text exceeds the column width with Aspose.Cells | How to detect overflowed cell text and enforce tooltip comments in an Excel file using Aspose.Cells .NET | Aspose.Cells example for verifying comment presence based on cell text truncation in a workbook | Programmatically validate that tooltips are attached only to truncated cells in Excel via Aspose.Cells | Determine when to add a comment to a cell by comparing string length to column width in Aspose.Cells C#
// Tags: Aspose.Cells detect cell text overflow | Aspose.Cells validate comment based on column width | Excel column width truncation check C# | Aspose.Cells iterate used range cells | tooltip presence conditional on cell overflow

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, iterates through all used cells, compares each cell's displayed string length to its column width (character units) to decide if the text is truncated, checks whether a comment (tooltip) exists, logs mismatches where a tooltip is present on a non‑truncated cell or missing on a truncated one, and finally saves the workbook.
class TooltipVerification
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Skip empty cells
                    if (cell.Value == null) continue;

                    // Get the cell's string value (displayed text)
                    string rawText = cell.StringValue;

                    // Get the column width in characters (Aspose.Cells returns width in characters)
                    double columnWidth = sheet.Cells.GetColumnWidth(col);

                    // Determine if the text would be truncated based on length vs column width
                    bool isTruncated = rawText.Length > columnWidth;

                    // Check if a comment (tooltip) is attached to the cell
                    bool hasTooltip = cell.Comment != null;

                    // Verify that tooltip appears only when text is truncated
                    if (hasTooltip && !isTruncated)
                    {
                        Console.WriteLine($"[Invalid] Tooltip found on non‑truncated cell {cell.Name}.");
                    }
                    else if (isTruncated && !hasTooltip)
                    {
                        Console.WriteLine($"[Warning] Cell {cell.Name} is truncated but has no tooltip.");
                    }
                    // Otherwise the condition is satisfied; no output needed
                }
            }

            // Save the workbook (lifecycle rule) – optional if modifications were made
            try
            {
                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
