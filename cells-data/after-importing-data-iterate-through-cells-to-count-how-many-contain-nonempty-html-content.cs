// Title: Count cells with non‑empty HTML content in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells count cells that contain HTML tags | How to find and count non‑empty HTML strings in Excel cells with Aspose.Cells | Iterate over used cells in a workbook and detect HTML markup using Aspose.Cells .NET | Example of counting cells with HTML content after stripping tags in Aspose.Cells | Detect HTML content in Excel worksheet cells using regular expressions and Aspose.Cells
// Tags: count html cells Aspose.Cells | detect html markup in Excel cells C# | iterate used cells Aspose.Cells .NET | regex html detection Aspose.Cells | non-empty html content counting Excel

using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;

// The sample loads an Excel file with Aspose.Cells, iterates through all used cells in the first worksheet, uses a compiled regular expression to identify HTML tags, strips the tags, and counts only those cells where the remaining text is non‑empty. It handles missing files and password‑protected workbooks, and outputs the total number of cells containing non‑empty HTML content.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook; if it is password‑protected without a password,
            // an exception will be thrown and handled in the catch block.
            workbook = new Workbook(inputPath);
        }
        catch (CellsException ex) when (ex.Message.Contains("password"))
        {
            Console.WriteLine("Error: The workbook is password protected and cannot be opened without a password.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            int htmlCellCount = 0;

            // Regex to detect HTML tags
            Regex htmlTagRegex = new Regex("<[^>]+>", RegexOptions.Compiled);

            // Iterate through all used cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Skip empty cells
                if (cell.Value == null) continue;

                string cellText = cell.Value.ToString();

                // Skip cells that are whitespace only
                if (string.IsNullOrWhiteSpace(cellText)) continue;

                // If the cell contains HTML tags
                if (htmlTagRegex.IsMatch(cellText))
                {
                    // Remove the tags and trim the remaining text
                    string stripped = htmlTagRegex.Replace(cellText, "").Trim();

                    // Count only if there is non‑empty content after stripping tags
                    if (!string.IsNullOrWhiteSpace(stripped))
                    {
                        htmlCellCount++;
                    }
                }
            }

            Console.WriteLine($"Number of cells with non‑empty HTML content: {htmlCellCount}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
