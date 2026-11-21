// Title: Export an Aspose.Cells workbook to HTML using a MemoryStream in C#
// Common Searches: C# Aspose.Cells save workbook as HTML to a MemoryStream | generate HTML from an Excel workbook in memory with Aspose.Cells | Aspose.Cells SaveFormat.Html stream example without writing a file | read HTML output from Aspose.Cells workbook via MemoryStream
// Tags: Aspose.Cells HTML stream export C# | save workbook to MemoryStream Aspose.Cells | in‑memory HTML conversion Aspose.Cells | SaveFormat.Html usage Aspose.Cells C# | C# memory stream workbook to HTML

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, adds sample data, and saves it directly to a MemoryStream as HTML using SaveFormat.Html. After resetting the stream position, the HTML can be read or processed in memory without creating a physical file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates a default workbook with one worksheet

        // Optionally, add some data to the worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Prepare a memory stream to hold the HTML output
        using (MemoryStream htmlStream = new MemoryStream())
        {
            // Save the workbook as HTML into the memory stream
            workbook.Save(htmlStream, SaveFormat.Html);

            // Reset the stream position to the beginning for further processing
            htmlStream.Position = 0;

            // Example: read the HTML content as a string (optional)
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                string htmlContent = reader.ReadToEnd();
                Console.WriteLine("Generated HTML length: " + htmlContent.Length);
                // Further in‑memory processing can be done here
            }
        }
    }
}
