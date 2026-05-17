using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ExportXmlUtf8Demo
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XmlMap found in the workbook.");
            return;
        }

        // Get the name of the first XML map
        string mapName = workbook.Worksheets.XmlMaps[0].Name;

        // Define the output XML file path
        string outputPath = "output.xml";

        // Export the XML map to a memory stream first
        using (MemoryStream memoryStream = new MemoryStream())
        {
            // ExportXml writes the XML data into the provided stream
            workbook.ExportXml(mapName, memoryStream);

            // Reset the stream position to read from the beginning
            memoryStream.Position = 0;

            // Read the XML content using UTF‑8 encoding
            string xmlContent = new StreamReader(memoryStream, Encoding.UTF8).ReadToEnd();

            // Write the XML content to the file with UTF‑8 encoding (including BOM)
            File.WriteAllText(outputPath, xmlContent, new UTF8Encoding(true));
        }

        Console.WriteLine($"XML exported successfully to '{outputPath}' with UTF‑8 encoding.");
    }
}