// Title: C# Example: Retrieve Cells Mapped to an XPath with Worksheet.XmlMapQuery in Aspose.Cells
// Description: Demonstrates how to import XML with namespaces into a workbook, obtain the generated XmlMap, define an XPath, and use Worksheet.XmlMapQuery to return the cell areas linked to that path. The sample prints each cell address and value, then saves the workbook.
// Keywords: Aspose.Cells | Worksheet.XmlMapQuery | C# XML map example | XPath cell query | import XML to Excel | mapped cell areas | XmlMap retrieval | Excel XML mapping .NET | cell address from XPath | Aspose.Cells .NET tutorial
// Common Searches: Worksheet.XmlMapQuery C# example | find cells by XPath Aspose.Cells | retrieve mapped cell ranges from XML map | Aspose.Cells query cells linked to XML node | how to use XmlMapQuery with namespaces
// Developer Intent: Locate worksheet cells that are associated with a specific XPath expression in an imported XML map.
// Use Cases: Validate or update values in cells that correspond to a particular XML element. | Generate an audit report of cell addresses linked to XML nodes for data lineage tracking. | Extract mapped cell values to feed downstream business logic or reporting pipelines. | Programmatically synchronize Excel data with an external XML source based on XPath mappings.
// AI Prompts: Show how to query multiple XPath expressions with Worksheet.XmlMapQuery and combine the results. | Provide code that automatically resolves XML namespace prefixes when calling XmlMapQuery. | Explain how to extend the example to map additional XML nodes and retrieve their cell ranges.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapQueryDemo
{
    // Demonstrates how to import XML with namespaces into a workbook, obtain the generated XmlMap, define an XPath, and use Worksheet.XmlMapQuery to return the cell areas linked to that path. The sample prints each cell address and value, then saves the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Sample XML data to be imported
            string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
                <ns1:Root xmlns:ns1='http://example.com'>
                    <ns1:Data>
                        <ns1:Item>Value1</ns1:Item>
                        <ns1:Item>Value2</ns1:Item>
                    </ns1:Data>
                </ns1:Root>";

            // Convert the XML string to a MemoryStream (required for ImportXml overload with Stream)
            using (MemoryStream xmlStream = new MemoryStream())
            {
                using (StreamWriter writer = new StreamWriter(xmlStream))
                {
                    writer.Write(xmlData);
                    writer.Flush();
                    xmlStream.Position = 0;

                    // Import XML into the first worksheet starting at cell A1 (lifecycle rule: load)
                    workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                }
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first (and only) XmlMap created by ImportXml
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the XPath expression to query mapped cells
            string xpath = "/ns1:Root/ns1:Data/ns1:Item";

            // Query the worksheet for cell areas linked to the specified XPath (feature rule)
            ArrayList cellAreas = worksheet.XmlMapQuery(xpath, xmlMap);

            // Output the results
            if (cellAreas.Count > 0)
            {
                Console.WriteLine($"Found {cellAreas.Count} cell area(s) mapped to path '{xpath}':");
                foreach (CellArea area in cellAreas)
                {
                    // For each area, display start row/column and the cell value
                    string cellAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                    Console.WriteLine($"- Cell {cellAddress}: {cellValue}");
                }
            }
            else
            {
                Console.WriteLine($"No cells are mapped to the path '{xpath}'.");
            }

            // Save the workbook to verify that the XML mapping persists (lifecycle rule: save)
            workbook.Save("XmlMapQueryResult.xlsx");
        }
    }
}
