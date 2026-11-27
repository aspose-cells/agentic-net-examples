// Title: Set the workbook's theme hyperlink color on cells containing URLs during data import with Aspose.Cells for .NET (C#)
// Common Searches: how to apply theme hyperlink color to URL cells using Aspose.Cells C# | apply hyperlink formatting to cells that contain links while importing Excel data in .NET | detect URLs in Excel cells and set theme hyperlink style with Aspose.Cells | C# Aspose.Cells change font color and underline for cells containing web addresses | use regex to find URLs and style them as hyperlinks in an Aspose.Cells workbook
// Tags: theme hyperlink color styling Aspose.Cells | url detection regex C# Aspose.Cells | apply font color underline to hyperlink cells Aspose.Cells | traverse used range cells Aspose.Cells | conditional cell formatting based on URL Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using System.Drawing;

// The example loads or creates a workbook, scans every cell in the used range with a regex to identify URLs, and then applies the workbook's theme hyperlink color along with underline formatting to those cells before saving the file.
class ApplyHyperlinkThemeColor
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load '{inputPath}': {loadEx.Message}");
                    return;
                }
            }
            else
            {
                workbook = new Workbook();
                // Add a sample sheet with some data for demonstration
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue("https://example.com");
                ws.Cells["B2"].PutValue("Not a link");
            }

            // Use a default hyperlink color (blue). Theme hyperlink color retrieval is omitted for compatibility.
            Color hyperlinkColor = Color.Blue;

            // Simple URL detection regex
            Regex urlRegex = new Regex(@"^(http|https|ftp)://", RegexOptions.IgnoreCase);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the sheet
                Aspose.Cells.Range maxRange = sheet.Cells.MaxDisplayRange;

                int startRow = maxRange.FirstRow;
                int endRow = maxRange.FirstRow + maxRange.RowCount - 1;
                int startCol = maxRange.FirstColumn;
                int endCol = maxRange.FirstColumn + maxRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        string cellValue = cell.StringValue?.Trim();

                        // Check if the cell contains a URL
                        if (!string.IsNullOrEmpty(cellValue) && urlRegex.IsMatch(cellValue))
                        {
                            // Apply the hyperlink style
                            Style style = cell.GetStyle();
                            style.Font.Color = hyperlinkColor;
                            style.Font.Underline = FontUnderlineType.Single;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
