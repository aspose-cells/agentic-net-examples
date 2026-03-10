using System;
using Aspose.Cells;

class RetrieveXmlMapRoot
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath); // workbook-load rule

        // Check if any XML maps are present
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Access the first XML map in the collection
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0]; // collection indexer

            // Get the root element name of the XML map
            string rootElementName = xmlMap.RootElementName;

            // Display the root element name
            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }

        // Optional: save the workbook if modifications were made
        // workbook.Save("output.xlsx", SaveFormat.Xlsx); // workbook-save rule
    }
}