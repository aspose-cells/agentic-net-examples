// Title: Omit unused named styles from HTML output when saving a workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, defines both a used and an unused style, applies only the used style, saves the workbook to HTML with HtmlSaveOptions.ExcludeUnusedStyles set to true, and verifies that the unused style's CSS is not present. | Show how to read the generated HTML file in C# and programmatically check for specific color definitions to confirm that only the applied style was exported. | Enhance the example with robust error handling and console messages that indicate whether each style was included or omitted in the HTML output.
// Common Searches: Aspose.Cells HTMLSaveOptions ExcludeUnusedStyles C# example | How to prevent unused Excel styles from being written to HTML with Aspose.Cells | Check if Aspose.Cells removes redundant CSS when exporting to HTML | C# code to verify that only applied cell styles appear in exported HTML | Exclude unused named styles from HTML using Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions exclude unused styles | C# export workbook to HTML without redundant CSS | verify used cell style in Aspose.Cells HTML output | remove unused named styles Aspose.Cells HTML export | HTML export style omission Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The program creates a workbook, defines a red unused style and a blue used style, applies the used style to cell A1, saves the workbook as HTML with ExcludeUnusedStyles enabled, then reads the HTML file to confirm that the CSS for the unused red style is absent while the CSS for the used blue style is present.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define an unused style (red font)
            Style unusedStyle = workbook.CreateStyle();
            unusedStyle.Font.Color = Color.Red;

            // Define a used style (blue font)
            Style usedStyle = workbook.CreateStyle();
            usedStyle.Font.Color = Color.Blue;

            // Apply the used style to cell A1
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["A1"].SetStyle(usedStyle);

            // Set HTML save options to exclude unused styles
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExcludeUnusedStyles = true
            };

            string htmlFile = "output.html";

            // Save workbook as HTML using the options
            workbook.Save(htmlFile, htmlOptions);

            // Verify that the unused style is omitted from the HTML output
            if (File.Exists(htmlFile))
            {
                string htmlContent = File.ReadAllText(htmlFile);

                // Look for CSS color definitions corresponding to the styles
                bool unusedStyleFound = htmlContent.Contains("color:#ff0000") || htmlContent.Contains("color:red");
                bool usedStyleFound = htmlContent.Contains("color:#0000ff") || htmlContent.Contains("color:blue");

                Console.WriteLine($"Unused style present in HTML: {unusedStyleFound}");
                Console.WriteLine($"Used style present in HTML: {usedStyleFound}");
            }
            else
            {
                Console.WriteLine($"HTML file '{htmlFile}' was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
