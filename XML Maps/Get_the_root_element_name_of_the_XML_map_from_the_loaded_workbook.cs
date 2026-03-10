using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (workbook-load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the workbook contains any XML maps
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Get the root element name of the XML map
            string rootElementName = xmlMap.RootElementName;

            // Display the result
            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}