// Title: Retrieve the root element name of the first XML map in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, verifies that XML maps exist, and prints the RootElementName of the first XmlMap. | Show a .NET example that loads a workbook, accesses the Worksheets.XmlMaps collection, and outputs the root element name of the first map.
// Common Searches: aspocells c# get root element name of first xml map in workbook | how to read XmlMap.RootElementName using Aspose.Cells .NET | example code for accessing XML maps collection in an Excel file with Aspose.Cells | retrieve root element name from Excel XML map via Aspose.Cells API
// Tags: Aspose.Cells XmlMap.RootElementName | C# read first XML map root element | Aspose.Cells load workbook XML maps | Excel XML map root element extraction | Aspose.Cells check XmlMaps collection

using Aspose.Cells;
using System;

// Demonstrates how to load an Excel workbook with Aspose.Cells, verify the presence of XML maps, obtain the first XmlMap object, read its RootElementName property, and display the name on the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Check that at least one XML map exists
        if (workbook.Worksheets.XmlMaps.Count > 0)
        {
            // Get the first XML map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Retrieve the root element name of the XML map
            string rootElementName = xmlMap.RootElementName;

            // Display the result
            Console.WriteLine("Root element name: " + rootElementName);
        }
        else
        {
            Console.WriteLine("No XML maps found in the workbook.");
        }
    }
}
