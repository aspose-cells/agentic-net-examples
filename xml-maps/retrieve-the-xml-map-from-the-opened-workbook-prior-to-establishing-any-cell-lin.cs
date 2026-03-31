using System;
using Aspose.Cells;

class RetrieveXmlMap
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the collection of XML maps in the workbook
        XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

        // Verify that at least one XML map is present
        if (xmlMaps.Count > 0)
        {
            // Retrieve the first XML map (or any specific one by index)
            XmlMap xmlMap = xmlMaps[0];

            // Output basic information about the retrieved XML map
            Console.WriteLine("XML Map Name: " + xmlMap.Name);
            Console.WriteLine("Root Element Name: " + xmlMap.RootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }

        // No cell links are created; this code only retrieves the XML map.
    }
}