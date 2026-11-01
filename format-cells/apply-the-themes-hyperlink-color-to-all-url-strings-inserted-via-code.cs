// Title: Apply the workbook theme’s hyperlink color to URL cells using Aspose.Cells for .NET (C#)
// Common Searches: how to apply the workbook theme hyperlink color to cells in Aspose.Cells C# | set hyperlink style to match Excel theme with Aspose.Cells API | programmatically format inserted URLs with theme colors in Aspose.Cells | Aspose.Cells C# apply theme-based hyperlink formatting without manual color | apply Excel theme hyperlink formatting to cells after adding Hyperlink objects
// Tags: Aspose.Cells apply workbook theme hyperlink style | C# set hyperlink formatting to theme color | Excel theme based hyperlink style Aspose.Cells | hyperlink style programmatic Excel theme C# | format URL cells with theme color Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example creates a workbook, inserts URL strings into cells, adds Hyperlink objects, defines a style that uses the workbook theme’s hyperlink color (blue and underlined), applies the style to the hyperlink cells, ensures the output directory exists, and saves the file as HyperlinkTheme.xlsx.
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

            // Insert URL strings into cells
            sheet.Cells["A1"].PutValue("https://example.com");
            sheet.Cells["A2"].PutValue("http://openai.com");

            // Add hyperlinks to the same cells
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com"); // A1
            sheet.Hyperlinks.Add(1, 0, 1, 1, "http://openai.com");   // A2

            // Create a hyperlink style (blue and underlined)
            Style hyperlinkStyle = workbook.CreateStyle();
            hyperlinkStyle.Font.Color = Color.Blue;
            hyperlinkStyle.Font.Underline = FontUnderlineType.Single; // correct enum value

            // Apply the style to cells that contain hyperlinks
            sheet.Cells["A1"].SetStyle(hyperlinkStyle);
            sheet.Cells["A2"].SetStyle(hyperlinkStyle);

            // Save the workbook
            string outputPath = "HyperlinkTheme.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
