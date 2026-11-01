// Title: How to exclude unused named styles from HTML output using Aspose.Cells HtmlSaveOptions.ExcludeUnusedStyles in C#
// Common Searches: C# Aspose.Cells save workbook to HTML without unused styles | example of HtmlSaveOptions.ExcludeUnusedStyles property in .NET | verify that red italic style is omitted from generated HTML with Aspose.Cells | how to filter out unused named styles when exporting Excel to HTML in C# | Aspose.Cells HTML export style cleanup using ExcludeUnusedStyles option
// Tags: Aspose.Cells HTML export omit redundant styles | HtmlSaveOptions ExcludeUnusedStyles usage in .NET | named style removal Aspose.Cells | HTML CSS cleanup with Aspose.Cells | C# test for omitted styles in HTML output | Excel to HTML style pruning Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, defines a used blue‑bold style and an unused red‑italic style, applies the used style to cell A1, and saves the workbook as HTML with HtmlSaveOptions.ExcludeUnusedStyles set to true. It then reads the generated HTML, confirms that the styled text appears while the CSS for the unused red style is absent, and prints verification results.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create a style that will be used
            Style usedStyle = workbook.CreateStyle();
            usedStyle.Font.Color = Color.Blue;
            usedStyle.Font.IsBold = true;

            // Create a style that will NOT be used
            Style unusedStyle = workbook.CreateStyle();
            unusedStyle.Font.Color = Color.Red;
            unusedStyle.Font.IsItalic = true;

            // Apply the used style to a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Styled Text");
            cell.SetStyle(usedStyle);

            // Prepare HTML save options with ExcludeUnusedStyles enabled
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExcludeUnusedStyles = true // Omit unused styles
            };

            // Define output HTML file path
            string htmlPath = "output.html";

            // Save the workbook as HTML
            workbook.Save(htmlPath, saveOptions);

            // Ensure the file was created before reading
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("HTML output file was not created.", htmlPath);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // Verify that the styled text is present
            bool textPresent = htmlContent.Contains("Styled Text");
            // Verify that the unused style's color (red) does not appear
            bool unusedStylePresent = htmlContent.Contains("color:#FF0000"); // red color CSS

            // Output verification results
            Console.WriteLine("Styled text present in HTML: " + textPresent);
            Console.WriteLine("Unused style (red color) present in HTML: " + unusedStylePresent);

            // Simple assertion
            if (textPresent && !unusedStylePresent)
            {
                Console.WriteLine("Verification succeeded: Unused styles are omitted.");
            }
            else
            {
                Console.WriteLine("Verification failed: Unexpected style presence.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
