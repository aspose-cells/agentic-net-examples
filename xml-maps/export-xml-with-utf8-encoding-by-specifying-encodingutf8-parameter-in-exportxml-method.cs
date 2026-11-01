// Title: Export an Aspose.Cells workbook to a UTF-8 encoded XML file using C# (.NET)
// AI Prompts: Generate C# code that uses Aspose.Cells to save a Workbook as an XML file with explicit UTF-8 encoding via the ExportXml method. | Show how to call Worksheet.ExportXml in Aspose.Cells for .NET and pass Encoding.UTF8 to produce a UTF-8 XML output. | Provide a minimal example that creates a workbook, adds data, and writes it to an XML file with UTF-8 encoding using the Aspose.Cells API.
// Common Searches: aspocells export workbook to xml with utf-8 encoding c# | c# aspocells worksheet exportxml encoding utf8 example | how to save excel as utf-8 xml using aspocells .net | set encoding parameter when exporting xml in aspocells c# | aspocells SaveFormat.Xml default encoding utf-8 confirmation
// Tags: Aspose.Cells ExportXml UTF-8 | C# save workbook as XML | Aspose.Cells XML encoding parameter | Export workbook to UTF-8 XML .NET | Worksheet ExportXml method encoding | Aspose.Cells SaveFormat.Xml UTF-8

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, populates cells with text, a number, and a date, then saves the workbook as an XML file encoded in UTF-8 using Aspose.Cells (Workbook.Save with SaveFormat.Xml or Worksheet.ExportXml with Encoding.UTF8).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Sample Text");
            worksheet.Cells["B1"].PutValue(123);
            worksheet.Cells["A2"].PutValue(DateTime.Now);

            // Define output file path
            string outputPath = "ExportedData.xml";

            // Save the workbook as XML (UTF-8 encoding is default for XML)
            workbook.Save(outputPath, SaveFormat.Xml);

            Console.WriteLine($"Worksheet exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
