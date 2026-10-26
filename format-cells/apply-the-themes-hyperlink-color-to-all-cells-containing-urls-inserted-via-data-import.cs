// Title: Apply the workbook’s theme hyperlink color to cells containing URLs during data import with Aspose.Cells for .NET
// Common Searches: how to automatically style URL strings as hyperlinks in Aspose.Cells C# | set theme hyperlink color for cells that contain URLs when importing data with Aspose.Cells | add hyperlink formatting to detected URLs in an Excel file using Aspose.Cells .NET | apply underline and theme color to cells with URLs during workbook processing Aspose.Cells | detect and hyperlink URLs in imported Excel data using Aspose.Cells C#
// Tags: theme hyperlink color formatting Aspose.Cells | URL detection in Excel cells C# | add cell hyperlink programmatically Aspose.Cells | style string cells as hyperlink Aspose.Cells | iterate used range workbook Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, scans every used string cell for URLs, applies the theme's hyperlink color with underline, inserts a hyperlink for each URL, and saves the updated file.
    class Program
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
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Helper method to determine whether a string looks like a URL
                bool IsUrl(string text)
                {
                    if (string.IsNullOrWhiteSpace(text))
                        return false;

                    return Uri.TryCreate(text, UriKind.Absolute, out Uri uriResult) &&
                           (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
                }

                // Use a standard hyperlink color (theme hyperlink color not directly exposed)
                Color hyperlinkColor = Color.Blue;

                // Get the used range of the worksheet (fully qualified to avoid ambiguity with System.Range)
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Iterate over the used cells
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;

                            // If the cell contains a URL, apply hyperlink formatting
                            if (IsUrl(text))
                            {
                                // Apply hyperlink color and underline
                                Style style = cell.GetStyle();
                                style.Font.Color = hyperlinkColor;
                                style.Font.Underline = FontUnderlineType.Single;
                                cell.SetStyle(style);

                                // Add a hyperlink (Aspose.Cells will replace an existing one if present)
                                sheet.Hyperlinks.Add(row, col, 1, 1, text);
                            }
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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
