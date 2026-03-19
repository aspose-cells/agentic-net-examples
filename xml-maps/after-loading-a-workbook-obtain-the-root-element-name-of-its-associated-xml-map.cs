using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Check if the workbook contains any XML maps
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map in the collection
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Get the root element name of the XML map
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