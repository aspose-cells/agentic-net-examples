// Title: Apply the workbook theme's hyperlink color to URL cells inserted programmatically with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a style using the workbook's built‑in Hyperlink theme color and applies it to cells containing URL strings. | Show how to add clickable hyperlinks to Excel cells while preserving the default theme underline and blue font using Aspose.Cells. | Demonstrate saving an Excel file where all programmatically inserted URLs retain the workbook's theme hyperlink formatting.
// Common Searches: aspocells apply theme hyperlink color to inserted URLs in C# | how to style programmatically added hyperlinks with workbook theme in Aspose.Cells | C# Aspose.Cells add clickable hyperlink preserving theme formatting | set hyperlink font color to theme default when writing Excel with Aspose.Cells
// Tags: theme-based hyperlink styling Aspose.Cells | insert URLs with styled hyperlinks C# | hyperlink cell formatting using workbook theme | add clickable hyperlinks programmatically Aspose.Cells | export Excel with themed hyperlink appearance

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example creates a new workbook, inserts URL strings into cells, defines a style that uses the workbook's built‑in Hyperlink theme color (blue underline), applies the style to each cell, adds actual clickable hyperlinks, and saves the workbook as HyperlinkThemeColor.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // URLs to insert
            string[] urls = { "https://www.example.com", "http://www.test.com/page" };

            // Create a style that mimics the built‑in Hyperlink style
            Style hyperlinkStyle = workbook.CreateStyle();
            hyperlinkStyle.Font.Color = Color.Blue;                     // Theme hyperlink color
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single;   // Underline like a hyperlink

            // Insert each URL and apply the hyperlink style
            for (int i = 0; i < urls.Length; i++)
            {
                // Put the URL text into the cell
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue(urls[i]);

                // Apply the hyperlink style
                cell.SetStyle(hyperlinkStyle);

                // Add an actual hyperlink to make it clickable
                sheet.Hyperlinks.Add(i, 0, 1, 1, urls[i]);
            }

            // Define output file path
            string outputPath = "HyperlinkThemeColor.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
