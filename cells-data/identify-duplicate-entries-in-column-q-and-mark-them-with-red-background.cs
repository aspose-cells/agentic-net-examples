// Title: Highlight duplicate values in column Q of an Excel workbook with a red background using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# highlight duplicate cells in column Q with red fill | How to mark repeated values in Excel column Q using Aspose.Cells | C# code to find and style duplicate entries in a specific Excel column with Aspose.Cells | Apply red background to duplicate cells in column Q using Aspose.Cells library | Detect and color duplicate values in column Q of an .xlsx file with Aspose.Cells C#
// Tags: Aspose.Cells duplicate detection in Excel column | apply red fill style Aspose.Cells C# | highlight duplicate cells column Q Aspose.Cells | Excel column duplicate marking Aspose.Cells | C# style cells based on value Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, counts occurrences of non‑empty values in column Q (index 16), creates a solid red fill style, and applies this style to every cell whose value appears more than once, then saves the updated file while handling missing files and password‑protected workbooks.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook; if the file is password‑protected, an exception will be thrown
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Column Q (zero‑based index 16)
            const int columnIndex = 16;

            // Determine the last used row in the worksheet
            int lastRow = sheet.Cells.MaxDataRow;

            // First pass: count occurrences of each cell value in column Q
            var valueCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = sheet.Cells[row, columnIndex];
                string text = cell.StringValue?.Trim() ?? string.Empty;

                if (string.IsNullOrEmpty(text))
                    continue; // Skip empty cells

                if (valueCounts.ContainsKey(text))
                    valueCounts[text]++;
                else
                    valueCounts[text] = 1;
            }

            // Prepare a style with a red background
            Style redStyle = workbook.CreateStyle();
            redStyle.ForegroundColor = Color.Red;
            redStyle.Pattern = BackgroundType.Solid;

            // Second pass: apply the red style to duplicate entries
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = sheet.Cells[row, columnIndex];
                string text = cell.StringValue?.Trim() ?? string.Empty;

                if (string.IsNullOrEmpty(text))
                    continue;

                if (valueCounts.TryGetValue(text, out int count) && count > 1)
                {
                    cell.SetStyle(redStyle);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Output saved to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected workbook or other Aspose.Cells specific errors
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password‑protected. Provide a valid password to load it.");
            }
            else
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
