// Title: Import XML data with an XSD map into an Excel workbook using Aspose.Cells Workbook.ImportXml (C#)
// Description: Demonstrates how to add an XML map from a schema.xsd file to a new Aspose.Cells workbook, import matching data.xml into Sheet1 at cell A1 with Workbook.ImportXml, and save the result as an .xlsx file.
// Keywords: Aspose.Cells | C# | Workbook.ImportXml | XML map | XSD schema | import XML to Excel | .NET Excel automation | Excel XML import example | global
// Common Searches: Aspose.Cells add XML map from XSD | Workbook.ImportXml C# example | Import XML into Excel using Aspose.Cells | How to use XSD schema with Aspose.Cells | C# code to map XML to Excel worksheet
// Developer Intent: Add an XSD‑based XML map to a workbook and import XML data that conforms to the schema into a specific worksheet.
// Use Cases: Create a fresh workbook, attach an XSD schema as an XML map, and load XML data into the first sheet for reporting. | Validate that the XML map was added correctly by reading its Name and RootElementName before importing. | Reuse the same XML map to import multiple XML files into different sheets or cell ranges within one workbook.
// AI Prompts: Write C# code that adds an XML map from a .xsd file to an Aspose.Cells workbook and imports a matching .xml file using Workbook.ImportXml. | Explain how to verify XmlMap properties (Name, RootElementName) after adding the map and before importing data. | Provide step‑by‑step instructions for importing XML data into a specific worksheet cell with an XSD schema in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapImportDemo
{
    // Demonstrates how to add an XML map from a schema.xsd file to a new Aspose.Cells workbook, import matching data.xml into Sheet1 at cell A1 with Workbook.ImportXml, and save the result as an .xlsx file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Add an XML map to the workbook using the XSD schema file (XmlMapCollection.Add rule)
            // The XSD file defines the structure of the XML data that will be imported.
            int mapIndex = workbook.Worksheets.XmlMaps.Add("schema.xsd");

            // Optionally retrieve the added XmlMap (not required for ImportXml, but useful for verification)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            Console.WriteLine($"Added XML map with name: {xmlMap.Name}, root element: {xmlMap.RootElementName}");

            // Import XML data into the first worksheet starting at cell A1 (Workbook.ImportXml rule)
            // The XML file must conform to the previously added XSD schema.
            workbook.ImportXml("data.xml", "Sheet1", 0, 0);

            // Save the workbook to an Excel file (save rule)
            workbook.Save("XmlMapImportedWorkbook.xlsx");

            Console.WriteLine("Workbook saved successfully with imported XML data.");
        }
    }
}
