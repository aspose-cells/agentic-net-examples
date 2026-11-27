// Title: How to retrieve cell areas mapped to an XML XPath using Worksheet.XmlMapQuery in Aspose.Cells for C#
// Common Searches: Aspose.Cells C# retrieve cells linked to an XML map XPath | Worksheet.XmlMapQuery example with namespace prefixes | Get cell ranges for a specific XML path in a workbook using Aspose.Cells | C# code to list cell addresses mapped to an XML element with Aspose.Cells | How to query mapped cells after importing XML into an Aspose.Cells worksheet
// Tags: Worksheet.XmlMapQuery cell area extraction | Aspose.Cells XML namespace mapping | C# import XML into worksheet with Aspose.Cells | Enumerate mapped cell addresses using XmlMapQuery | CellArea list from XML path query

using System;
using System.Collections;
using Aspose.Cells;

// The sample creates a Workbook, imports an XML string containing a namespace (which automatically creates an XmlMap), defines an XPath with the namespace prefix, and calls Worksheet.XmlMapQuery to obtain an ArrayList of CellArea objects representing the cells mapped to that path. It then iterates the results, prints each cell's address and value, and saves the workbook as 'XmlMapQueryResult.xlsx'.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample XML containing a namespace and two items
        string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns1:Root xmlns:ns1='http://example.com'>
    <ns1:Data>
        <ns1:Item>Value1</ns1:Item>
        <ns1:Item>Value2</ns1:Item>
    </ns1:Data>
</ns1:Root>";

        // Import the XML into the worksheet.
        // This creates an XmlMap automatically and populates cells starting at A1.
        workbook.ImportXml(xml, "Sheet1", 0, 0);

        // Verify that an XmlMap was created
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML maps were created during import.");
            return;
        }

        // Retrieve the first (and only) XmlMap
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML path we want to query.
        // The path must include the namespace prefix defined in the XML.
        string path = "/ns1:Root/ns1:Data/ns1:Item";

        // Query the worksheet for cell areas linked to the specified XML path.
        // XmlMapQuery returns an ArrayList of CellArea objects.
        ArrayList cellAreas = worksheet.XmlMapQuery(path, xmlMap);

        // Output the query results
        if (cellAreas == null || cellAreas.Count == 0)
        {
            Console.WriteLine($"No cells are mapped to the XML path '{path}'.");
        }
        else
        {
            Console.WriteLine($"Found {cellAreas.Count} cell area(s) mapped to the XML path '{path}':");
            for (int i = 0; i < cellAreas.Count; i++)
            {
                CellArea area = (CellArea)cellAreas[i];
                // Each Item maps to a single cell, so start == end.
                string address = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                Console.WriteLine($"  Area {i + 1}: {address} (Row {area.StartRow}, Column {area.StartColumn})");
                Console.WriteLine($"    Cell value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
            }
        }

        // Save the workbook so you can inspect the mapping visually if desired.
        workbook.Save("XmlMapQueryResult.xlsx");
    }
}
