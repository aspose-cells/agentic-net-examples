using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHtmlBorderDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a style with a double border (many browsers do not support this directly)
                Style doubleBorderStyle = workbook.CreateStyle();
                doubleBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.TopBorder].Color = Color.DarkBlue;
                doubleBorderStyle.Borders[BorderType.BottomBorder].Color = Color.DarkBlue;
                doubleBorderStyle.Borders[BorderType.LeftBorder].Color = Color.DarkBlue;
                doubleBorderStyle.Borders[BorderType.RightBorder].Color = Color.DarkBlue;

                // Apply the style to a range of cells so the border is visible in the HTML output
                Aspose.Cells.Range range = sheet.Cells.CreateRange("B2:D4");
                range.ApplyStyle(doubleBorderStyle, new StyleFlag { Borders = true });

                // Fill the range with sample data
                for (int row = 1; row <= 3; row++)
                {
                    for (int col = 1; col <= 3; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row}C{col}");
                    }
                }

                // Configure HTML save options to export a similar border style when the original is not supported
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportSimilarBorderStyle = true, // Enable fallback border rendering
                    IsBorderCollapsed = false        // Keep borders separate for clearer CSS inspection
                };

                // Define output HTML file path
                string outputPath = Path.Combine(Environment.CurrentDirectory, "FallbackBorderDemo.html");

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"HTML file saved to: {outputPath}");

                // Read the generated HTML and output CSS border definitions for manual inspection
                if (File.Exists(outputPath))
                {
                    string[] htmlLines = File.ReadAllLines(outputPath);
                    Console.WriteLine("\n--- Extracted CSS Border Rules ---");
                    foreach (string line in htmlLines)
                    {
                        // Look for CSS rules that contain the word 'border'
                        if (line.IndexOf("border", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            Console.WriteLine(line.Trim());
                        }
                    }
                    Console.WriteLine("--- End of CSS Extraction ---");
                }
                else
                {
                    Console.WriteLine("Failed to locate the generated HTML file.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}