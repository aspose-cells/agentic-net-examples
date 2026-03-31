using System;
using Aspose.Cells;

class ExportXmlUsingMap
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook wb = new Workbook("input.xlsx");

        // Check if any XML maps are present
        if (wb.Worksheets.XmlMaps.Count > 0)
        {
            // Retrieve the first XML map in the collection
            XmlMap xmlMap = wb.Worksheets.XmlMaps[0];

            // Export the XML data linked by this map to a file
            wb.ExportXml(xmlMap.Name, "output.xml");
            Console.WriteLine("XML exported successfully to output.xml");
        }
        else
        {
            Console.WriteLine("No XmlMap found in the workbook");
        }
    }
}