// Title: Namespace‑Aware XPath Query on Worksheet Cells with Worksheet.XmlMapQuery (C# Aspose.Cells)
// Description: Demonstrates how to import XML with a namespace into an Aspose.Cells workbook, retrieve the generated XmlMap, and use Worksheet.XmlMapQuery with a prefixed XPath (e.g., "/ns:Root/ns:Data/ns:Item") to obtain the CellArea objects that correspond to the mapped XML nodes. The sample prints each cell's address and value and optionally saves the workbook.
// Keywords: Aspose.Cells | Worksheet.XmlMapQuery | namespace aware XPath | C# XML map query | XmlMapQuery example | import XML with namespace | CellArea collection | C# Aspose.Cells tutorial | XPath to cells | XML map to worksheet
// Common Searches: Aspose.Cells XmlMapQuery namespace prefix example | C# query worksheet cells using XPath with namespace | How to map XML nodes to cells in Aspose.Cells | Retrieve cell addresses from XML map Aspose.Cells | Namespace‑aware XPath in Aspose.Cells workbook
// Developer Intent: Locate and read worksheet cells that are linked to XML nodes by executing a namespace‑aware XPath query on an imported XML map.
// Use Cases: Extract all cells representing <ns:Item> elements after importing a namespaced XML document. | Validate that specific namespaced XML nodes are correctly mapped to worksheet cells by checking their coordinates and values. | Generate a report of cell addresses for a given XML schema namespace to support data migration or auditing.
// AI Prompts: Provide a C# code snippet that uses Worksheet.XmlMapQuery with a custom namespace prefix to query an imported XML map and returns the matching CellArea objects. | Explain how to handle multiple XmlMaps in a workbook and select the appropriate map when performing a namespace‑aware XPath query. | Show how to iterate over the CellArea collection returned by XmlMapQuery and copy the matched cell values to another worksheet.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to import XML with a namespace into an Aspose.Cells workbook, retrieve the generated XmlMap, and use Worksheet.XmlMapQuery with a prefixed XPath (e.g., "/ns:Root/ns:Data/ns:Item") to obtain the CellArea objects that correspond to the mapped XML nodes. The sample prints each cell's address and value and optionally saves the workbook.
    public class NamespaceAwareXmlMapQueryDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample XML containing a namespace declaration
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns:Root xmlns:ns='http://example.com/schema'>
    <ns:Data>
        <ns:Item>Value1</ns:Item>
        <ns:Item>Value2</ns:Item>
    </ns:Data>
</ns:Root>";

                // Import the XML into the worksheet; this also creates an XML map
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // Retrieve the generated XML map (the first one in the collection)
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Define a namespace‑aware XPath using the prefix defined in the XML (ns)
                string xpath = "/ns:Root/ns:Data/ns:Item";

                // Query the worksheet for cell areas linked to the specified XPath
                ArrayList cellAreas = worksheet.XmlMapQuery(xpath, xmlMap);

                // Output the results
                if (cellAreas.Count > 0)
                {
                    foreach (CellArea area in cellAreas)
                    {
                        // For each matched area, display its start cell coordinates and value
                        Console.WriteLine($"Found data at Row {area.StartRow + 1}, Column {area.StartColumn + 1}");
                        Console.WriteLine($"Cell value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
                    }
                }
                else
                {
                    Console.WriteLine("No cells were mapped to the specified XPath.");
                }

                // Save the workbook (optional, demonstrates saving)
                workbook.Save("NamespaceAwareXmlMapQueryDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            NamespaceAwareXmlMapQueryDemo.Run();
        }
    }
}
