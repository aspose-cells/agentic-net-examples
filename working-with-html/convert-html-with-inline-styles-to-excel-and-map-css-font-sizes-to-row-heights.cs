// Title: Convert HTML containing inline CSS to an Excel workbook and set row heights based on CSS font sizes using Aspose.Cells for .NET
// Common Searches: asp.net convert html with inline styles to xlsx preserving formatting | set excel row height from css font-size when importing html with Aspose.Cells | how to map html font size to excel row height in C# | load html file into Aspose.Cells workbook and adjust row heights programmatically | convert html to excel while keeping inline css styling using Aspose.Cells
// Tags: Aspose.Cells HTML import row height adjustment | CSS font-size to Excel row height mapping | inline CSS styling preservation Aspose.Cells | C# calculate row height from cell font size | save workbook as XLSX after HTML conversion

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithRowHeight
{
    // The example loads an HTML file with inline CSS into an Aspose.Cells Workbook using HtmlLoadOptions, scans each cell for non‑default font sizes, computes a row height (font size × 1.2 points) keeping the maximum per row, applies these heights to the worksheet rows, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // -----------------------------------------------------------------
            // 1. Define input and output paths.
            // -----------------------------------------------------------------
            string htmlPath = "input.html";          // Path to the source HTML file.
            string excelPath = "output.xlsx";        // Path where the Excel file will be saved.

            // Validate input file existence.
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            try
            {
                // -----------------------------------------------------------------
                // 2. Load the HTML into a Workbook using Aspose.Cells.
                // -----------------------------------------------------------------
                // HtmlLoadOptions allows us to keep the original formatting.
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();

                // Use the constructor that accepts a file name and load options.
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // -----------------------------------------------------------------
                // 3. Iterate through all worksheets and cells to map font size → row height.
                // -----------------------------------------------------------------
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Dictionary to keep the maximum required height per row.
                    Dictionary<int, double> maxRowHeight = new Dictionary<int, double>();

                    // Get the used range to limit iteration.
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell == null) continue;

                            Style style = cell.GetStyle();
                            // Font.Size is expressed in points. Zero means default (not set).
                            if (style != null && style.Font != null && style.Font.Size > 0)
                            {
                                double fontSize = style.Font.Size; // points
                                // Simple conversion factor: 1 point font → 1.2 points row height.
                                double desiredHeight = fontSize * 1.2;

                                // Keep the greatest height required for the row.
                                if (maxRowHeight.TryGetValue(row, out double existingHeight))
                                {
                                    if (desiredHeight > existingHeight)
                                        maxRowHeight[row] = desiredHeight;
                                }
                                else
                                {
                                    maxRowHeight[row] = desiredHeight;
                                }
                            }
                        }
                    }

                    // Apply the calculated heights.
                    foreach (var kvp in maxRowHeight)
                    {
                        int rowIndex = kvp.Key;
                        double height = kvp.Value;
                        // Row height is also in points.
                        sheet.Cells.Rows[rowIndex].Height = height;
                    }
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook as an Excel file.
                // -----------------------------------------------------------------
                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(excelPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(excelPath, SaveFormat.Xlsx);
                Console.WriteLine($"Conversion completed successfully. Excel file saved to '{excelPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during conversion:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
