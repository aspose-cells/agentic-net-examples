// Title: Attach an XML Schema Map to an Excel Workbook with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load an existing .xlsx file into an Aspose.Cells Workbook, add an XSD schema as an XML map via the Worksheets.XmlMaps collection, optionally assign a friendly name, and save the workbook with the map attached.
// Keywords: Aspose.Cells | XML map | XSD schema | C# | .NET | load Excel workbook | add XML map | Worksheets.XmlMaps | save workbook with schema | Excel XML integration
// Common Searches: Aspose.Cells add XML map to workbook C# | how to attach XSD to Excel using Aspose.Cells | load .xlsx and add schema map .NET | C# code for XML schema map in Excel | Aspose.Cells Worksheets.XmlMaps example
// Developer Intent: Add an XML schema map to an existing Excel workbook and persist the change.
// Use Cases: Prepare a template workbook for XML data exchange by embedding an XSD map. | Give an XML map a readable name for later import/export operations. | Distribute a workbook that already contains the required XML structure for downstream processing.
// AI Prompts: Show how to add multiple XML schema maps to a single workbook with Aspose.Cells. | Provide C# code that imports XML data into a workbook using a previously attached XML map. | Explain how to retrieve the index or name of an XML map after it has been added.

using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    // Demonstrates how to load an existing .xlsx file into an Aspose.Cells Workbook, add an XSD schema as an XML map via the Worksheets.XmlMaps collection, optionally assign a friendly name, and save the workbook with the map attached.
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook from a .xlsx file
            Workbook workbook = new Workbook("input.xlsx");

            // Add an XML schema map to the workbook
            // The schema file path can be a local file or a URL
            int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

            // Optionally set a friendly name for the map
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MySchemaMap";

            // Save the workbook with the attached XML map
            workbook.Save("output.xlsx");

            Console.WriteLine("Workbook loaded and XML schema map attached successfully.");
        }
    }
}
