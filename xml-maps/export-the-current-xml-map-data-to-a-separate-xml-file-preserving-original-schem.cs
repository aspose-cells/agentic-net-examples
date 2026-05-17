using System;
using Aspose.Cells;

class ExportXmlMapDemo
{
    static void Main()
    {
        // Load the workbook containing the XML map
        Workbook workbook = new Workbook("input.xlsx");

        // Verify that at least one XML map exists
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Retrieve the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Export the XML data using the map's name, preserving the original schema
            workbook.ExportXml(xmlMap.Name, "exported.xml");

            Console.WriteLine("XML map exported successfully to 'exported.xml'.");
        }
        else
        {
            Console.WriteLine("No XML map found in the workbook.");
        }
    }
}