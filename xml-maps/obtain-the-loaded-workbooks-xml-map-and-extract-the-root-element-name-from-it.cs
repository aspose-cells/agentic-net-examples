using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the workbook contains any XML maps
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Retrieve the first XmlMap from the collection
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Get the root element name of the XML map
            string rootElementName = xmlMap.RootElementName;

            // Output the root element name
            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}