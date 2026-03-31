using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapRootElementDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains an XML map
            string workbookPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(workbookPath);

            // Check if the workbook has any XML maps
            if (workbook.Worksheets.XmlMaps.Count > 0)
            {
                // Retrieve the first XML map in the collection
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
}