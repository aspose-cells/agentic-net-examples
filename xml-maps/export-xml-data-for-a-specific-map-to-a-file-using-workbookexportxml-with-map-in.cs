using System;
using Aspose.Cells;

class ExportXmlByMapIndex
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook wb = new Workbook("input.xlsx");

        // Index of the XML map to export (0‑based)
        int mapIndex = 0; // adjust as needed

        // Verify that the requested map exists
        if (wb.Worksheets.XmlMaps.Count > mapIndex)
        {
            // Retrieve the XmlMap by index
            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];

            // Export the XML data using the map's name
            wb.ExportXml(xmlMap.Name, "exported.xml");

            Console.WriteLine($"XML exported successfully using map index {mapIndex} to 'exported.xml'.");
        }
        else
        {
            Console.WriteLine($"No XmlMap found at index {mapIndex}.");
        }
    }
}