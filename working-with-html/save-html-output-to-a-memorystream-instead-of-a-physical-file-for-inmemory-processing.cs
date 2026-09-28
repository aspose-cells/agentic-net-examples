// Title: Export an Aspose.Cells Workbook to HTML using a MemoryStream in C#
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook, fills it with data, and saves it directly to a MemoryStream in HTML format. | Show how to reset the MemoryStream after saving HTML and retrieve the HTML string for further processing. | Refactor the example into a reusable method that returns the HTML MemoryStream from a given Workbook.
// Common Searches: aspocells save workbook as html to memory stream c# example | how to generate html from excel in memory using Aspose.Cells | c# read html output from Aspose.Cells workbook without writing a file | asp.net core return excel as html stream using Aspose.Cells
// Tags: Aspose.Cells HTML export to MemoryStream | C# in‑memory Excel to HTML conversion | HTML output from Aspose.Cells using MemoryStream | MemoryStream handling for Aspose.Cells output | File‑system‑free Excel HTML generation

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating an Aspose.Cells Workbook, populating cells, and saving it as HTML directly into a MemoryStream, then resetting the stream and reading the HTML string, enabling in‑memory Excel‑to‑HTML conversion without touching the file system.
class HtmlToMemoryStreamExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Prepare a memory stream to hold the HTML output
        using (MemoryStream htmlStream = new MemoryStream())
        {
            // Save the workbook as HTML into the memory stream
            workbook.Save(htmlStream, SaveFormat.Html);

            // Reset the stream position to the beginning for reading
            htmlStream.Position = 0;

            // Example: read the HTML content as a string (optional)
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                string htmlContent = reader.ReadToEnd();
                Console.WriteLine("Generated HTML:");
                Console.WriteLine(htmlContent);
            }

            // At this point, htmlStream contains the HTML representation of the workbook
            // It can be returned, sent over a network, or processed further without touching the file system
        }
    }
}
