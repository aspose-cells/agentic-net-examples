using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the workbook that contains an XML map
        string inputPath = "input.xlsx";

        // Load the workbook (lifecycle rule: workbook-load)
        Workbook workbook = new Workbook(inputPath);

        // Check if any XML maps are present
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Get the root element name of the map
            string rootElementName = xmlMap.RootElementName;

            // Output the result
            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}