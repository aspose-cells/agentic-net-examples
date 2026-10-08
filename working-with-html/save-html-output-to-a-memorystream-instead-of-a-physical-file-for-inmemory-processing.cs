// Title: Export an Aspose.Cells workbook to HTML using a MemoryStream in C#
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, populates it, and saves the HTML output directly into a MemoryStream. | Demonstrate how to read the HTML string from the MemoryStream after calling Workbook.Save and return it from a method. | Adapt the example to stream the generated HTML as an HTTP response in an ASP.NET Core controller.
// Common Searches: C# Aspose.Cells save workbook as HTML into a MemoryStream | how to get HTML string from Aspose.Cells workbook without creating a file | Aspose.Cells export Excel to HTML stream for web API | in‑memory conversion of Excel worksheet to HTML using Aspose.Cells | read HTML output from Aspose.Cells MemoryStream in .NET
// Tags: Aspose.Cells HTML export to MemoryStream | C# workbook to HTML stream | in‑memory Excel to HTML conversion | Aspose.Cells save as HTML without file | read HTML from MemoryStream Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new Aspose.Cells Workbook, adds data to the first worksheet, saves the workbook as HTML directly into a MemoryStream, resets the stream position, reads the HTML content as a string, and writes it to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and add some data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Prepare a MemoryStream to receive the HTML output
        using (MemoryStream htmlStream = new MemoryStream())
        {
            // Save the workbook as HTML into the memory stream
            workbook.Save(htmlStream, SaveFormat.Html);

            // Reset the stream position to the beginning for reading
            htmlStream.Position = 0;

            // Example: read the HTML content from the stream
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                string htmlContent = reader.ReadToEnd();
                Console.WriteLine(htmlContent);
            }
        }
    }
}
