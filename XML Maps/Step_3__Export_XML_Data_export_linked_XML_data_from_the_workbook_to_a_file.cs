using System;
using Aspose.Cells;

class ExportXmlDemo
{
    static void Main()
    {
        // Load the workbook that contains linked XML data
        Workbook workbook = new Workbook("input.xlsx"); // workbook-load

        // Verify that the workbook has at least one XML map
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Retrieve the name of the first XML map
            string mapName = workbook.Worksheets.XmlMaps[0].Name;

            // Export the XML data linked by the map to a file
            workbook.ExportXml(mapName, "exported.xml");

            Console.WriteLine("XML exported successfully to exported.xml");
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}