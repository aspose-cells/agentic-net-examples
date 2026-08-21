// Title: Apply Theme Hyperlink Color to URL Cells After Data Import with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, imports a DataTable that contains plain text and URLs, detects strings that start with http:// or https://, adds a hyperlink to each matching cell, and applies the workbook's theme hyperlink formatting (blue, underlined). The file is then saved as an Excel workbook.
// Keywords: Aspose.Cells C# hyperlink theme | apply theme hyperlink color .NET | detect URLs in worksheet Aspose | import DataTable Excel hyperlink | format hyperlinks after data import | Excel theme hyperlink style C# | Aspose.Cells add hyperlink programmatically
// Common Searches: how to style hyperlinks with workbook theme using Aspose.Cells | C# add hyperlink to cells containing URLs Aspose | apply default hyperlink color after importing data Excel | Aspose.Cells detect and convert URL strings to hyperlinks | set theme hyperlink formatting programmatically
// Developer Intent: Automatically convert URL strings in imported data to clickable hyperlinks and style them with the workbook’s default theme color.
// Use Cases: Convert a contacts list’s website column into clickable, theme‑styled hyperlinks after bulk import. | Generate a report where any cell containing an http/https string becomes a hyperlink that matches the workbook’s default appearance. | Prepare an export that scans imported text fields for URLs, adds hyperlinks, and applies the theme’s hyperlink style before saving.
// AI Prompts: Write C# code with Aspose.Cells that scans all used cells, adds a hyperlink to strings beginning with http:// or https://, and applies the workbook’s theme hyperlink style. | Provide an Aspose.Cells snippet that imports a DataTable and automatically formats any URL cells using the default theme color without hard‑coding the color value. | Explain how to retrieve the theme’s hyperlink color from a workbook and apply it to cells after adding hyperlinks in Aspose.Cells for .NET.

using System;
using System.Data;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace ApplyHyperlinkThemeColor
{
    // This example creates a workbook, imports a DataTable that contains plain text and URLs, detects strings that start with http:// or https://, adds a hyperlink to each matching cell, and applies the workbook's theme hyperlink formatting (blue, underlined). The file is then saved as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Simulate data import with URLs
                DataTable dt = new DataTable();
                dt.Columns.Add("Name");
                dt.Columns.Add("Website");
                dt.Rows.Add("Aspose", "https://www.aspose.com");
                dt.Rows.Add("Google", "http://www.google.com");
                dt.Rows.Add("NoLink", "Sample Text");

                // Import the DataTable into the worksheet (compatible with all versions)
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        sheet.Cells[r, c].PutValue(dt.Rows[r][c]?.ToString() ?? string.Empty);
                    }
                }

                // Create a hyperlink style (blue and underlined)
                Style hyperlinkStyle = workbook.CreateStyle();
                hyperlinkStyle.Font.Color = Color.Blue;
                hyperlinkStyle.Font.Underline = FontUnderlineType.Single;

                // Iterate through used cells to find URLs and apply hyperlinks with the style
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue.Trim();

                            // Simple URL detection (http or https)
                            if (text.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                                text.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                            {
                                // Add hyperlink to the cell (single cell range)
                                sheet.Hyperlinks.Add(row, col, 1, 1, text);

                                // Apply the hyperlink style to the cell
                                cell.SetStyle(hyperlinkStyle);
                            }
                        }
                    }
                }

                // Define output file path
                string outputPath = "OutputWithHyperlinkTheme.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
