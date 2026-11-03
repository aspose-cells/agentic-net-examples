// Title: C# example: Load HTML into an Aspose.Cells workbook while ignoring <img> tags and verifying cell content
// Common Searches: c# load html into workbook without importing images using Aspose.Cells | how to prevent img tags from being added as pictures when converting html to excel | verify that cell A1 contains paragraph text after html import with Aspose.Cells | remove pictures after loading html string into Aspose.Cells workbook c# example
// Tags: html load options ignore images aspocells | aspocells remove pictures after html import | verify cell value after html to excel conversion | excel export without images c# | load html string to workbook aspocells

using Aspose.Cells;
using System;
using System.IO;

// This C# sample loads an HTML string containing a paragraph and an <img> tag into an Aspose.Cells Workbook using HtmlLoadOptions, clears any pictures that might have been imported, confirms that the picture count is zero, checks that cell A1 holds the expected text "Hello World", and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // HTML content that includes a paragraph and an image tag
            string html = "<html><body><p>Hello World</p><img src='https://example.com/image.png' alt='test'></body></html>";

            // Load options (no specific image‑ignoring flag available in this version)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML into a workbook from a memory stream
            using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html)))
            {
                Workbook workbook = new Workbook(stream, loadOptions);

                // Remove any pictures that may have been imported
                if (workbook.Worksheets[0].Pictures.Count > 0)
                {
                    workbook.Worksheets[0].Pictures.Clear();
                }

                // Verify that no pictures were added
                int pictureCount = workbook.Worksheets[0].Pictures.Count;
                Console.WriteLine($"Picture count (expected 0): {pictureCount}");

                // Verify that the text from the paragraph is placed in a cell (A1 by default)
                string cellValue = workbook.Worksheets[0].Cells["A1"].StringValue;
                Console.WriteLine($"Cell A1 value (expected 'Hello World'): {cellValue}");

                // Save the workbook
                string outputPath = "Result.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
