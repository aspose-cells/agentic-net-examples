// Title: Convert an HTML table with embedded base64 images to an Excel workbook, inserting each image into the matching cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that reads an HTML string containing <img> tags with base64 data, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, and saves the result as an .xlsx file. | Explain how to ensure each image from the HTML table is placed into the exact worksheet cell that originally held the <img> element while keeping the table structure intact. | Provide a sample that loads the HTML from a file path, supports both base64‑encoded images and external image URLs, and exports the content to Excel using Aspose.Cells.
// Common Searches: asp.net convert html with base64 images to excel using aspose.cells c# | c# load html table containing pictures into workbook cells with Aspose.Cells | how to preserve image positions when exporting html to xlsx with Aspose.Cells | aspose.cells htmlloadoptions handling of embedded images in excel export | convert html data:image src to excel cells c# example
// Tags: html to xlsx conversion using Aspose.Cells | Aspose.Cells HtmlLoadOptions image handling | embed base64 pictures into Excel cells C# | preserve table layout during html to excel export | load external image URLs into workbook with Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// // This example loads an HTML string that contains a table with base64‑encoded <img> tags into an Aspose.Cells Workbook using HtmlLoadOptions, then saves the workbook as an .xlsx file, preserving the table structure and placing each image into its corresponding cell.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Input HTML string containing a table with embedded images (base64 or URLs)
        string html = @"<table>
<tr><td>Item 1</td><td><img src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA...'/></td></tr>
<tr><td>Item 2</td><td><img src='data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQ...'/></td></tr>
</table>";

        try
        {
            // Convert the HTML string to a memory stream
            byte[] htmlBytes = Encoding.UTF8.GetBytes(html);
            using (MemoryStream htmlStream = new MemoryStream(htmlBytes))
            {
                // Load HTML directly into a workbook; Aspose.Cells handles tables and images (including base64)
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Define output file name
                string outputPath = "ConvertedFromHtml.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Conversion completed. File saved as {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
