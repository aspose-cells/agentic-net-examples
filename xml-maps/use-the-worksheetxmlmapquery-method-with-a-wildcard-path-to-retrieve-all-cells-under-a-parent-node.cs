// Title: Worksheet.XmlMapQuery with wildcard to fetch all child cells from an XML map (C#)
// Description: Demonstrates how to import XML into a workbook, create an XmlMap, define a wildcard path ("/Root/Parent/*"), and use Worksheet.XmlMapQuery to return CellArea objects for every cell linked to any child element under a specific XML parent. The example iterates the results, prints cell addresses and values, and saves the workbook.
// Keywords: Aspose.Cells | Worksheet.XmlMapQuery | wildcard path | XML map query | C# | retrieve mapped cells | asterisk wildcard | CellArea enumeration | ImportXml | XML to Excel mapping
// Common Searches: Worksheet.XmlMapQuery wildcard example | Aspose.Cells get cells for all child nodes | C# query XML map with * path | How to list cells linked to XML parent in Aspose.Cells | XmlMapQuery return multiple CellArea objects
// Developer Intent: Find a concise way to query a worksheet for every cell that is mapped to any child element of a given XML parent node using a wildcard path.
// Use Cases: Extract values of all dynamic child elements after importing an XML document. | Validate that each XML child node has a corresponding cell in the worksheet. | Generate a dynamic list of cell addresses for further processing when element names are unknown.
// AI Prompts: Write a C# snippet that uses Worksheet.XmlMapQuery with "*" to list cell addresses for all nodes under a specified XML parent. | Explain the role of the asterisk wildcard in XmlMapQuery and show safe iteration over the returned CellArea collection. | Show how to filter XmlMapQuery results to include only cells containing non‑empty string values.

using System;
using System.Collections;
using Aspose.Cells;

// Demonstrates how to import XML into a workbook, create an XmlMap, define a wildcard path ("/Root/Parent/*"), and use Worksheet.XmlMapQuery to return CellArea objects for every cell linked to any child element under a specific XML parent. The example iterates the results, prints cell addresses and values, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Sample XML containing a parent node with multiple child elements
        string xml = @"<?xml version='1.0' encoding='UTF-8'?>
            <Root>
                <Parent>
                    <Child>Value1</Child>
                    <Child>Value2</Child>
                    <Child>Value3</Child>
                </Parent>
                <Other>Ignore</Other>
            </Root>";

        // Import the XML into the first worksheet starting at cell A1
        wb.ImportXml(xml, "Sheet1", 0, 0);

        // Retrieve the first XmlMap created by ImportXml
        if (wb.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML map found in the workbook.");
            return;
        }
        XmlMap xmlMap = wb.Worksheets.XmlMaps[0];

        // Define a wildcard path that selects all direct children of the Parent node
        // The asterisk (*) matches any child element under /Root/Parent
        string wildcardPath = "/Root/Parent/*";

        // Query the worksheet for cell areas linked to the wildcard path
        Worksheet ws = wb.Worksheets[0];
        ArrayList cellAreas = ws.XmlMapQuery(wildcardPath, xmlMap);

        // Iterate through the returned CellArea objects and display cell information
        foreach (CellArea area in cellAreas)
        {
            // Each CellArea typically represents a single cell (StartRow == EndRow, StartColumn == EndColumn)
            Cell cell = ws.Cells[area.StartRow, area.StartColumn];
            Console.WriteLine($"Cell {cell.Name} maps to path '{wildcardPath}' with value: {cell.StringValue}");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        wb.Save("XmlMapQueryWildcardDemo.xlsx");
    }
}
