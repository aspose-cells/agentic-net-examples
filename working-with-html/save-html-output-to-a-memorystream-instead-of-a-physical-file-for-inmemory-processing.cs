// Title: Save Aspose.Cells Workbook as HTML to a MemoryStream in C# (.NET)
// Description: Demonstrates how to generate HTML from an Aspose.Cells workbook directly into a MemoryStream, read the markup as a string, and avoid creating a temporary file. Ideal for web APIs, email bodies, or any scenario that requires in‑memory HTML processing.
// Keywords: Aspose.Cells HTML MemoryStream | C# export workbook to HTML stream | in‑memory HTML generation .NET | Aspose.Cells SaveFormat.Html stream | convert workbook to HTML without file
// Common Searches: Aspose.Cells save workbook as HTML to MemoryStream | C# generate HTML from Excel without writing file | How to use MemoryStream with Aspose.Cells Save method | In‑memory HTML export Aspose.Cells .NET
// Developer Intent: Create HTML output from a workbook directly in memory for further processing or transmission.
// Use Cases: Return HTML from an ASP.NET Core controller as a response. | Embed generated markup in an email template without disk I/O. | Store HTML in a database or pass it to another service.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as HTML into a MemoryStream and returns the HTML string. | Show how to send the MemoryStream HTML result from an ASP.NET Core endpoint using FileContentResult. | Explain the benefits of using a custom MemoryStream factory when exporting HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to generate HTML from an Aspose.Cells workbook directly into a MemoryStream, read the markup as a string, and avoid creating a temporary file. Ideal for web APIs, email bodies, or any scenario that requires in‑memory HTML processing.
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Aspose.Cells");

        // Create a MemoryStream using the provided factory method
        CustomImplementationFactory factory = new CustomImplementationFactory();
        MemoryStream htmlStream = factory.CreateMemoryStream();

        // Save the workbook as HTML directly into the memory stream
        workbook.Save(htmlStream, SaveFormat.Html);

        // Reset the stream position to read the generated HTML
        htmlStream.Position = 0;
        string htmlContent = new StreamReader(htmlStream).ReadToEnd();

        // Output the HTML content (for demonstration purposes)
        Console.WriteLine(htmlContent);
    }
}
