// Title: Extract and count dominant foreground and background cell colors from an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file using Aspose.Cells, iterates every worksheet and cell, and records each unique System.Drawing.Color from the cell's foreground and background styles together with its occurrence count. | Write a method that orders the collected colors by frequency and prints the ARGB components along with the number of times each color appears. | Add validation to confirm the Excel file exists before processing and display a clear error message if the path is invalid.
// Common Searches: how to list all cell background colors in an Excel file using Aspose.Cells C# | C# Aspose.Cells count how many times each color is used in a workbook | extract dominant ARGB colors from Excel worksheets with Aspose.Cells .NET | enumerate foreground and background style colors across all sheets in .xlsx using Aspose.Cells | retrieve color palette statistics from Excel workbook programmatically in C#
// Tags: extract cell style colors Aspose.Cells .NET | count foreground and background colors in Excel workbook | dominant ARGB color analysis using Aspose.Cells | iterate worksheets to collect cell colors C# | track color frequencies Aspose.Cells

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, scans every cell in each worksheet, captures non‑transparent foreground and background System.Drawing.Color values, tallies their occurrences in a dictionary, sorts the colors by usage descending, and outputs each ARGB value with its count.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = @"C:\Data\SampleWorkbook.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" does not exist.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Dictionary to hold color usage count across all sheets
            Dictionary<Color, int> colorUsage = new Dictionary<Color, int>();

            // Iterate through each worksheet
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan each cell in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell == null) continue;

                        // Get the cell style
                        Style style = cell.GetStyle();

                        // Collect foreground color if set
                        Color fg = style.ForegroundColor;
                        if (!fg.IsEmpty && fg.A != 0)
                            IncrementColorCount(colorUsage, fg);

                        // Collect background color if set
                        Color bg = style.BackgroundColor;
                        if (!bg.IsEmpty && bg.A != 0)
                            IncrementColorCount(colorUsage, bg);
                    }
                }
            }

            // Output the collected colors sorted by usage (descending)
            Console.WriteLine("Dominant colors across all sheets:");
            foreach (var kvp in SortedByUsage(colorUsage))
            {
                Color c = kvp.Key;
                int count = kvp.Value;
                Console.WriteLine($"Color ARGB({c.A},{c.R},{c.G},{c.B}) - Used {count} times");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Helper to increment color count in the dictionary
    private static void IncrementColorCount(Dictionary<Color, int> dict, Color color)
    {
        if (dict.ContainsKey(color))
            dict[color]++;
        else
            dict[color] = 1;
    }

    // Helper to sort dictionary by value descending
    private static IEnumerable<KeyValuePair<Color, int>> SortedByUsage(Dictionary<Color, int> dict)
    {
        var list = new List<KeyValuePair<Color, int>>(dict);
        list.Sort((a, b) => b.Value.CompareTo(a.Value));
        return list;
    }
}
