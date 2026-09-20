// Title: Set a workbook’s default font to “Courier New” and confirm HTML fallback using HtmlSaveOptions in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that assigns "Courier New" as the workbook’s DefaultStyle font, applies a non‑existent font to a cell, saves the workbook to HTML with HtmlSaveOptions, and then parses the resulting HTML to ensure the CSS contains a "font-family: \"Courier New\"" fallback. | Adapt the example to use a different custom default font and a new output path while still detecting whether the generated HTML correctly falls back to the specified default when the cell’s font cannot be resolved.
// Common Searches: Aspose.Cells C# set default workbook font for HTML export | how to verify font fallback in Aspose.Cells HTML output | HtmlSaveOptions default font not applied in generated HTML Aspose.Cells | C# Aspose.Cells export to HTML with custom fallback font | check font-family declaration in Aspose.Cells HTML file
// Tags: Aspose.Cells default workbook font C# | HtmlSaveOptions HTML export fallback font | verify generated HTML font-family Aspose.Cells | apply nonexistent font to cell Aspose.Cells | set default style font Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook, sets its DefaultStyle font to "Courier New", forces a missing font on a cell, saves the workbook as HTML using HtmlSaveOptions, and then reads the output file to confirm that the generated CSS includes a "Courier New" font-family fallback.
class HtmlDefaultFontDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set a default font for the workbook (used when a specified font is missing)
            workbook.DefaultStyle.Font.Name = "Courier New";
            workbook.DefaultStyle.Font.Size = 12;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some text to a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Sample text with missing font");

            // Apply a font that does not exist on the system to force fallback
            Style style = cell.GetStyle();
            style.Font.Name = "NonExistentFont";
            style.Font.Size = 12;
            cell.SetStyle(style);

            // Configure HtmlSaveOptions (no DefaultFont property needed)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Define output HTML file path
            string htmlPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);

            // Verify fallback behavior by inspecting the generated HTML
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                if (htmlContent.Contains("font-family: 'Courier New'") ||
                    htmlContent.Contains("font-family: \"Courier New\""))
                {
                    Console.WriteLine("Fallback to default font 'Courier New' verified.");
                }
                else
                {
                    Console.WriteLine("Default font fallback not applied as expected.");
                }
            }
            else
            {
                Console.WriteLine($"Failed to generate HTML file at path: {htmlPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
