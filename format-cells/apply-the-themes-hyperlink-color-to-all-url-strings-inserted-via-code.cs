// Title: Apply Workbook Theme Hyperlink Color to Programmatically Added URLs with Aspose.Cells for .NET
// Description: Creates a workbook, inserts URL hyperlinks into cells, extracts the theme's hyperlink color from the default style, builds a single style with that color and underline, and applies it to every cell that contains a hyperlink before saving the file.
// Keywords: Aspose.Cells C# hyperlink theme color | set hyperlink font color programmatically | apply default theme style to Excel links | format hyperlink cells Aspose.Cells | .NET Excel hyperlink styling
// Common Searches: Aspose.Cells apply theme hyperlink color | C# set hyperlink style to workbook theme | how to format hyperlink cells with default theme in Aspose.Cells | programmatically change Excel hyperlink color .NET
// Developer Intent: Use Aspose.Cells to style all programmatically added hyperlink cells with the workbook's theme hyperlink color.
// Use Cases: Generate Excel reports where every inserted URL automatically matches the document's theme color. | Batch‑add hyperlinks to a sheet and ensure consistent visual formatting without manual styling. | Refresh hyperlink appearance after changing the workbook theme to keep the UI cohesive.
// AI Prompts: Show how to retrieve the theme hyperlink color from a workbook's default style and apply it to hyperlink cells using Aspose.Cells for .NET. | Refactor the code to apply a single style object to all hyperlink ranges without iterating each cell. | Explain how to keep hyperlink formatting synchronized with theme changes after hyperlinks have been created.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts URL hyperlinks into cells, extracts the theme's hyperlink color from the default style, builds a single style with that color and underline, and applies it to every cell that contains a hyperlink before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add several hyperlinks (URL strings) to the worksheet
                worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");
                worksheet.Hyperlinks.Add("B2", 1, 1, "https://docs.aspose.com");
                worksheet.Hyperlinks.Add("C3", 1, 1, "https://github.com/aspose");

                // Retrieve the theme's hyperlink color from the default style
                Style hyperlinkStyle = workbook.DefaultStyle;
                // Use the default font color (theme hyperlink color) and underline
                hyperlinkStyle.Font.Color = workbook.DefaultStyle.Font.Color;
                hyperlinkStyle.Font.Underline = FontUnderlineType.Single;

                // Apply the style to each cell that contains a hyperlink
                foreach (Hyperlink link in worksheet.Hyperlinks)
                {
                    // The hyperlink's Area gives the range it occupies
                    int firstRow = link.Area.StartRow;
                    int firstColumn = link.Area.StartColumn;
                    int totalRows = link.Area.EndRow - link.Area.StartRow + 1;
                    int totalColumns = link.Area.EndColumn - link.Area.StartColumn + 1;

                    // Apply the style to every cell in the hyperlink range
                    for (int r = firstRow; r < firstRow + totalRows; r++)
                    {
                        for (int c = firstColumn; c < firstColumn + totalColumns; c++)
                        {
                            Cell cell = worksheet.Cells[r, c];
                            cell.SetStyle(hyperlinkStyle);
                        }
                    }
                }

                // Determine output path and ensure its directory exists
                string outputPath = "HyperlinksWithThemeColor.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
