using System;
using Aspose.Cells;

class RetrieveXmlMapRootElement
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // workbook-load rule

        // Check if any XML maps exist
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Get the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Retrieve the root element name
            string rootElementName = xmlMap.RootElementName;

            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}