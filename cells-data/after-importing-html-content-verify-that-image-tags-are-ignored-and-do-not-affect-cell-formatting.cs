// Title: Load HTML into an Aspose.Cells workbook in C# while ignoring <img> tags and preserving paragraph font styles
// AI Prompts: Read an HTML string with Aspose.Cells, apply HtmlLoadOptions, and ensure that no Picture objects are created in the worksheet. | Inspect the cells generated from <p> elements to confirm that bold and red font attributes are correctly applied. | Save the validated workbook to an XLSX file after confirming image tags were ignored.
// Common Searches: Aspose.Cells C# load HTML string without importing images | How to keep bold and colored text from HTML when converting to Excel with Aspose.Cells | Verify that picture collection is empty after loading HTML in Aspose.Cells | Using MemoryStream to load HTML into Aspose.Cells workbook C# example | Check cell style properties after HTML import Aspose.Cells C#
// Tags: html load ignore images Aspose.Cells | preserve paragraph font styling Aspose.Cells | validate picture count after html import C# | memory stream html to workbook Aspose.Cells | cell style verification Aspose.Cells C#

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Text;

// The C# program loads an HTML snippet containing bold text, an <img> tag, and red text into an Aspose.Cells Workbook via a MemoryStream, confirms that no pictures were imported, validates that the first cell is bold and the second cell has red font, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        // HTML containing an image tag and formatted text
        string html = @"
            <html>
                <body>
                    <p style='font-weight:bold;'>Bold Text</p>
                    <img src='https://example.com/image.png' />
                    <p style='color:red;'>Red Text</p>
                </body>
            </html>";

        Workbook workbook = null;

        try
        {
            // Load the HTML content into a workbook using a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream ms = new MemoryStream(htmlBytes))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(ms, loadOptions);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading HTML into workbook: {ex.Message}");
            return;
        }

        // Reference to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Verify that image tags were ignored (no pictures should be present)
        if (sheet.Pictures.Count == 0)
        {
            Console.WriteLine("Image tags were ignored as expected.");
        }
        else
        {
            Console.WriteLine($"Unexpected pictures count: {sheet.Pictures.Count}");
        }

        // Verify that cell formatting from HTML is applied correctly
        // Aspose.Cells maps each <p> element to a separate cell
        Cell cellBold = sheet.Cells[0, 0]; // First <p> (Bold Text)
        Cell cellRed = sheet.Cells[1, 0];  // Second <p> (Red Text)

        // Check if the first cell is bold
        bool isBold = cellBold.GetStyle().Font.IsBold;

        // Check if the second cell has red font color
        bool isRed = cellRed.GetStyle().Font.Color.ToArgb() == Color.Red.ToArgb();

        Console.WriteLine($"Cell A1 bold: {isBold}");
        Console.WriteLine($"Cell A2 red font: {isRed}");

        // Save the workbook
        try
        {
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
