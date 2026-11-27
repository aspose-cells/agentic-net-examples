// Title: Set workbook theme hyperlink color for programmatically added URL cells using Aspose.Cells for .NET
// Common Searches: how to apply the workbook theme's hyperlink color to cells in Aspose.Cells C# | Aspose.Cells retrieve ThemeColorType.Hyperlink and use it for hyperlink styling | C# add URLs to Excel and format them with the theme hyperlink color using Aspose.Cells | create a style with theme hyperlink color and underline in Aspose.Cells .NET | save Excel file with themed hyperlink formatting in Aspose.Cells
// Tags: theme hyperlink color retrieval Aspose.Cells | hyperlink style creation .NET | programmatic Excel hyperlink formatting C# | use workbook theme for cell formatting Aspose.Cells | save workbook with styled hyperlinks Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, obtains the hyperlink color defined in the workbook's theme (falling back to blue if unavailable), builds a style with that color and underline, inserts a list of URL strings into cells, adds matching hyperlinks, applies the themed style to each cell, and saves the workbook to a file.
class HyperlinkThemeColorExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // URLs to insert as hyperlinks
            string[] urls = { "https://example.com", "https://aspose.com", "https://github.com" };

            // Retrieve the hyperlink color defined in the workbook's theme (fallback to blue if unavailable)
            Color themeHyperlinkColor = Color.Blue;
            try
            {
                // Aspose.Cells provides GetThemeColor to obtain theme colors.
                themeHyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);
            }
            catch
            {
                // Keep the fallback color if any exception occurs.
            }

            // Create a style that uses the theme's hyperlink color and underline
            Style hyperlinkStyle = workbook.CreateStyle();
            hyperlinkStyle.Font.Color = themeHyperlinkColor;
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single;

            // Insert URL text, create hyperlinks, and apply the style
            for (int i = 0; i < urls.Length; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue(urls[i]); // display the URL text

                // Add a hyperlink that covers the same cell
                sheet.Hyperlinks.Add(i, 0, i, 0, urls[i]);

                // Apply the hyperlink style to the cell
                cell.SetStyle(hyperlinkStyle);
            }

            // Define output file path
            string outputPath = "HyperlinkThemeColor.xlsx";

            // Ensure the directory exists (handle cases where outputPath has no directory component)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
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
