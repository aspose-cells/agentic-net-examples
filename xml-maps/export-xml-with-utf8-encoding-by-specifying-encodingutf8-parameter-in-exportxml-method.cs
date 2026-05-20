using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ExportXmlUtf8Demo
{
    static void Main()
    {
        // Load the workbook that contains an XML map
        Workbook workbook = new Workbook("Book1.xlsx");

        // Verify that at least one XML map exists
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML map found in the workbook.");
            return;
        }

        // Get the name of the first XML map
        string xmlMapName = workbook.Worksheets.XmlMaps[0].Name;

        // Define the output XML file path
        string outputPath = "output_utf8.xml";

        // Create a FileStream and wrap it with a StreamWriter that uses UTF‑8 encoding
        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        using (StreamWriter writer = new StreamWriter(fileStream, Encoding.UTF8))
        {
            // Export the XML data using the underlying stream (UTF‑8 enforced by the writer)
            workbook.ExportXml(xmlMapName, writer.BaseStream);
        }

        Console.WriteLine($"XML exported successfully to '{outputPath}' with UTF‑8 encoding.");
    }
}