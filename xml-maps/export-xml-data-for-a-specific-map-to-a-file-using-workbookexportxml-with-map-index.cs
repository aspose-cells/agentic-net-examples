using System;
using Aspose.Cells;

class ExportXmlByMapIndex
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook wb = new Workbook("input.xlsx");

        // Index of the XML map to export (e.g., 0 for the first map)
        int mapIndex = 0;

        // Verify that the requested index exists
        if (mapIndex >= 0 && mapIndex < wb.Worksheets.XmlMaps.Count)
        {
            // Retrieve the XmlMap object by index
            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];

            // Export the XML data using the map's name
            wb.ExportXml(xmlMap.Name, "output.xml");

            Console.WriteLine("XML exported successfully to output.xml");
        }
        else
        {
            Console.WriteLine("The specified XML map index does not exist.");
        }
    }
}