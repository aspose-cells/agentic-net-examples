// Title: Highlight Excel cells that lack a leading quote using Aspose.Cells for .NET with a custom background color
// Common Searches: Aspose.Cells C# highlight cells without leading apostrophe | How to apply background color to cells where QuotePrefix is false in .NET | Conditional formatting based on QuotePrefix property using Aspose.Cells | Set style for cells that do not start with a single quote in C# Excel library | Iterate used range and color cells without quote prefix Aspose.Cells
// Tags: highlight cells without leading quote Aspose.Cells | conditional formatting based on QuotePrefix .NET | custom background style for Excel cells C# | iterate worksheet used range set style Aspose.Cells | apply style when QuotePrefix is false

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, creates a solid LightYellow style, scans the worksheet's used range, checks each cell's string value for a leading single quote, and applies the style only to cells where the QuotePrefix is false, then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Create a distinct background style (LightYellow)
                Style highlightStyle = workbook.CreateStyle();
                highlightStyle.ForegroundColor = Color.LightYellow;
                highlightStyle.Pattern = BackgroundType.Solid;

                // Get the used range of the worksheet (fully qualified to avoid ambiguity)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Iterate through each cell in the used range
                for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Determine if the cell has a leading quote (quote prefix)
                        bool hasQuotePrefix = false;
                        if (cell.Value != null && cell.Value is string s && s.StartsWith("'"))
                        {
                            hasQuotePrefix = true;
                        }

                        // Apply the highlight style only when the cell does NOT have a quote prefix
                        if (!hasQuotePrefix)
                        {
                            cell.SetStyle(highlightStyle);
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
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
