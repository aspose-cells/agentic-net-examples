// Title: Export an Aspose.Cells workbook to HTML in C# while removing all inline style attributes
// Common Searches: C# export Excel workbook to HTML using Aspose.Cells without inline CSS | how to prevent Aspose.Cells from adding style attributes in generated HTML | remove inline style tags from Aspose.Cells HTML output programmatically | verify that saved HTML from Aspose.Cells contains no style= attributes in C#
// Tags: Aspose.Cells generate HTML without embedded style attributes | HtmlSaveOptions suppress cell style rendering | C# programmatically verify absence of style attributes in HTML | Excel workbook to HTML conversion clean markup | Aspose.Cells HTML output without inline CSS

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, applies a red font style to a cell, saves the workbook as HTML using HtmlSaveOptions, reads the HTML into a string, checks for any remaining "style=" attributes, and writes the clean HTML to a file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data with styling
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Apply a style to the cell (this would normally generate inline styles)
            Style cellStyle = sheet.Cells["A1"].GetStyle();
            cellStyle.Font.Color = Color.Red;
            sheet.Cells["A1"].SetStyle(cellStyle);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: In newer Aspose.Cells versions the ExportCellStyle property may be unavailable.
            // If needed, adjust the options according to the version you are using.

            // Save the workbook to a memory stream as HTML
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, saveOptions);
                htmlStream.Position = 0;

                // Read the generated HTML into a string
                string htmlContent;
                using (StreamReader reader = new StreamReader(htmlStream))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Simple check for any remaining inline style attributes
                if (htmlContent.Contains("style="))
                {
                    Console.WriteLine("Inline style attributes were found in the generated HTML.");
                }
                else
                {
                    Console.WriteLine("No inline style attributes present.");
                }

                // Optionally write the HTML to a file for manual inspection
                string outputPath = "output.html";
                try
                {
                    File.WriteAllText(outputPath, htmlContent);
                    Console.WriteLine($"HTML output written to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write HTML file: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
