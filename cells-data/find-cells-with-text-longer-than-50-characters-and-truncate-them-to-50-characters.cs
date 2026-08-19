// Title: C# – Truncate Excel Cell Text Over 50 Characters with Aspose.Cells
// Description: Load an Excel workbook, scan each worksheet's used range, detect string cells longer than 50 characters, truncate the text to 50 characters while preserving formatting, and save the updated file.
// Keywords: Aspose.Cells truncate text | C# limit cell string length | Excel cell text cut off | Aspose.Cells used range iteration | save modified workbook Aspose.Cells | .NET Excel string trimming | cell value length check
// Common Searches: how to cut cell text to 50 characters using Aspose.Cells | Aspose.Cells truncate long strings in Excel | C# iterate used cells and modify values Aspose | remove characters beyond 50 in Excel cells .NET | save workbook after editing cell strings Aspose.Cells
// Developer Intent: Find and shorten any cell string that exceeds 50 characters, then write the changes back to a new Excel file.
// Use Cases: Process a large spreadsheet and ensure all textual entries fit a 50‑character limit. | Efficiently modify only populated cells by using MaxDataRow/MaxDataColumn. | Maintain original cell formatting while trimming overly long text.
// AI Prompts: Generate C# code with Aspose.Cells that truncates cell strings longer than a given length and saves the workbook. | Show how to loop through the used range of each worksheet and conditionally edit cell values. | Explain how to keep cell formatting intact when shortening text with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an Excel workbook, scan each worksheet's used range, detect string cells longer than 50 characters, truncate the text to 50 characters while preserving formatting, and save the updated file.
    public class TruncateLongTextDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (lifecycle: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to avoid scanning empty cells
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Loop through each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Process only cells that contain string values
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;

                            // If the text length exceeds 50 characters, truncate it
                            if (!string.IsNullOrEmpty(text) && text.Length > 50)
                            {
                                string truncated = text.Substring(0, 50);
                                cell.PutValue(truncated);
                            }
                        }
                    }
                }
            }

            // Save the modified workbook (lifecycle: save)
            string outputPath = "output_truncated.xlsx";

            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}
