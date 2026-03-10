using System;
using System.IO;
using Aspose.Cells;

class ExportXmlDemo
{
    static void Main()
    {
        string inputPath = "InputWithMap.xlsx";

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        // Load the workbook that contains an XML map
        Workbook workbook = new Workbook(inputPath);

        // Verify that at least one XML map is present
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Retrieve the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the output XML file path
            string outputPath = "ExportedData.xml";

            // Export the XML data linked by the map to the file
            workbook.ExportXml(xmlMap.Name, outputPath);

            Console.WriteLine($"XML exported successfully to {outputPath}");
        }
        else
        {
            Console.WriteLine("No XML map found in the workbook.");
        }
    }
}