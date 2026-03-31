using System;
using Aspose.Cells;

class RetrieveXmlMapRootElement
{
    static void Main()
    {
        // Load an existing workbook that contains XML maps
        Workbook workbook = new Workbook("input.xlsx");

        // Check that the workbook has at least one XML map
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map in the collection
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Retrieve the root element name of the XML map
            string rootElementName = xmlMap.RootElementName;

            // Display the root element name
            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}