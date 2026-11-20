// Title: Highlight cells containing the word “error” (case‑insensitive) with an orange background using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find cells with text 'error' ignoring case and apply orange fill | how to programmatically highlight error messages in Excel using Aspose.Cells .NET | search and format cells containing a specific keyword in a workbook with Aspose.Cells | apply conditional orange style to cells that contain the word error in C# | iterate through all worksheets and highlight matching strings using Aspose.Cells API
// Tags: case-insensitive text search Aspose.Cells | apply orange background style Aspose.Cells | highlight cells based on string value .NET | iterate worksheets cells Aspose.Cells | set cell style programmatically Excel .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates an orange solid‑fill style, iterates every worksheet and cell, checks each string cell for the word “error” without case sensitivity, applies the orange style to matching cells, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (no password supplied)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Create a style with orange background
            Style orangeStyle = workbook.CreateStyle();
            orangeStyle.ForegroundColor = Color.Orange;
            orangeStyle.Pattern = BackgroundType.Solid;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string cells containing "error"
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;
                            if (!string.IsNullOrEmpty(text) &&
                                text.IndexOf("error", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Apply the orange highlight
                                cell.SetStyle(orangeStyle);
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
