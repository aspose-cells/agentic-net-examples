using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class ExportXmlUtf8Demo
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("Book1.xlsx");

        // Check that the workbook contains at least one XML map
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML map found in the workbook.");
            return;
        }

        // Get the name of the first XML map
        string xmlMapName = workbook.Worksheets.XmlMaps[0].Name;

        // Define the output XML file path
        string outputPath = "output.xml";

        // Export the XML map to a file.
        // ExportXml writes the XML using UTF‑8 encoding by default.
        // The FileStream is used as the destination stream.
        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            workbook.ExportXml(xmlMapName, fileStream);
        }

        // Optional: read back the file using UTF‑8 to verify the content
        string xmlContent = File.ReadAllText(outputPath, Encoding.UTF8);
        Console.WriteLine("Exported XML (first 200 characters):");
        Console.WriteLine(xmlContent.Substring(0, Math.Min(200, xmlContent.Length)));
    }
}