// Title: Find and highlight duplicate values in a named range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells highlight duplicate cells in a named range | How to apply conditional formatting to duplicate values with Aspose.Cells .NET | Detect repeated entries in a specific Excel named range using Aspose.Cells | Programmatically style duplicate cells yellow in Excel with Aspose.Cells C# | Iterate over Aspose.Cells range to count and mark duplicate values
// Tags: named range duplicate value detection Aspose.Cells | yellow background style application Aspose.Cells C# | cell value counting within Aspose.Cells range | StyleFlag usage for formatting Aspose.Cells | programmatic duplicate highlighting in Excel .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

// The example loads an Excel workbook, retrieves the named range "MyRange", counts how many times each cell value occurs, creates a yellow background style, applies that style to every cell whose value appears more than once, and saves the updated file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Use fully qualified Aspose.Cells.Range to avoid ambiguity
            Aspose.Cells.Range range = namedRange.GetRange();

            // First pass: count occurrences of each cell value
            Dictionary<string, int> valueCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (Cell cell in range)
            {
                string val = cell.Value?.ToString() ?? string.Empty;
                if (valueCounts.ContainsKey(val))
                    valueCounts[val]++;
                else
                    valueCounts[val] = 1;
            }

            // Create a style to highlight duplicate cells
            Style duplicateStyle = workbook.CreateStyle();
            duplicateStyle.ForegroundColor = Color.Yellow;
            duplicateStyle.Pattern = BackgroundType.Solid;

            // Second pass: apply the highlight style to duplicate cells
            foreach (Cell cell in range)
            {
                string val = cell.Value?.ToString() ?? string.Empty;
                if (valueCounts.TryGetValue(val, out int count) && count > 1)
                {
                    StyleFlag flag = new StyleFlag { All = true };
                    cell.SetStyle(duplicateStyle, flag);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
