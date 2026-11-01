// Title: How to disable gridlines when exporting an Aspose.Cells workbook to HTML in C# and verify the output
// AI Prompts: Generate C# code that saves a Workbook to HTML with HtmlSaveOptions.ExportGridLines set to false and reads the file to confirm no gridline CSS is present. | Demonstrate how to use Aspose.Cells HtmlSaveOptions to omit gridlines in the HTML output and programmatically check that the generated HTML does not contain border styles.
// Common Searches: Aspose.Cells C# export workbook to HTML without gridlines | how to turn off gridlines in HTML output using Aspose.Cells HtmlSaveOptions | verify that exported HTML from Aspose.Cells does not contain gridline borders | C# code sample for disabling gridlines in Aspose.Cells HTML export | check HTML file for gridline CSS after saving with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions disable gridlines | ExportGridLines false C# | HTML export without borders Aspose.Cells | verify HTML gridline removal C# | Aspose.Cells workbook to HTML sample

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, populates it with sample data, configures HtmlSaveOptions with ExportGridLines = false (and optional Base64 image embedding), saves the workbook as HTML, reads the generated file, and checks for gridline‑related CSS to confirm that gridlines were omitted.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(20);

            // Configure HTML save options to ensure gridlines are not exported
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = false,          // Turn off gridlines in HTML
                ExportImagesAsBase64 = true       // Embed images as Base64 (optional)
            };

            // Define the output HTML file path
            string htmlFile = "ExportedWorkbook.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlFile, htmlOptions);

            // Verify that the HTML file was created
            if (!File.Exists(htmlFile))
            {
                Console.WriteLine($"Error: The file '{htmlFile}' was not created.");
                return;
            }

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlFile);

            // Simple check: look for typical gridline CSS patterns
            bool hasGridLines = htmlContent.Contains("gridlines") ||
                                htmlContent.Contains("border:1px solid") ||
                                htmlContent.Contains("border: thin solid");

            Console.WriteLine("Gridlines present in HTML output: " + hasGridLines);
            // Expected output: false
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
