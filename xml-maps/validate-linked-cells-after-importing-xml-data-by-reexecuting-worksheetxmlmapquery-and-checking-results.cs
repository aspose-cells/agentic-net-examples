// Title: Validate XML‑Linked Cells with Worksheet.XmlMapQuery in Aspose.Cells for .NET
// Description: This C# example shows how to import XML data from a memory stream into a workbook, retrieve the automatically created XmlMap, and re‑run Worksheet.XmlMapQuery for a specific XPath (e.g., "/Root/Data/Item"). It enumerates each linked cell area, prints the cell address and current value, handles the case of no matches, and optionally saves the workbook for visual inspection.
// Keywords: Aspose.Cells XmlMapQuery | C# import XML workbook | validate XML map links | .NET XML map verification | Worksheet.XmlMapQuery example | linked cell detection Aspose | XML data integrity check | Aspose.Cells memory stream import
// Common Searches: how to verify cells linked to an XML map in Aspose.Cells | Worksheet.XmlMapQuery usage after ImportXml | detect missing XML‑linked cells in a .NET workbook | Aspose.Cells validate imported XML data | C# query linked cells by XPath in Excel
// Developer Intent: Confirm that cells are correctly linked to a given XML path after importing XML into an Aspose.Cells workbook.
// Use Cases: Run a post‑import audit to ensure every <Item> element has a corresponding linked cell and list their addresses. | Trigger an alert when no cells are linked to a specified XPath, indicating a mapping or import issue. | Generate a quick report of linked cell counts for multiple XPaths to aid debugging of complex XML maps.
// AI Prompts: Create C# code that imports XML into an Aspose.Cells worksheet, then uses Worksheet.XmlMapQuery to return all linked cell addresses and values for a supplied XPath. | Write a reusable method that accepts a Workbook, an XmlMap, and an XPath, executes XmlMapQuery, and throws an exception if the result set is empty. | Develop a console application that imports XML, queries linked cells for several XPaths, and prints a summary table of each XPath with its linked cell count and sample values.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace XmlMapValidationDemo
{
    // This C# example shows how to import XML data from a memory stream into a workbook, retrieve the automatically created XmlMap, and re‑run Worksheet.XmlMapQuery for a specific XPath (e.g., "/Root/Data/Item"). It enumerates each linked cell area, prints the cell address and current value, handles the case of no matches, and optionally saves the workbook for visual inspection.
    class Program
    {
        static void Main()
        {
            // Sample XML data to be imported
            string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
                <Root>
                    <Data>
                        <Item>Value1</Item>
                        <Item>Value2</Item>
                    </Data>
                </Root>";

            // Convert the XML string to a MemoryStream (required by ImportXml overload)
            using (MemoryStream xmlStream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.Write(xmlData);
                writer.Flush();
                xmlStream.Position = 0; // Reset stream position before reading

                // Create a new workbook and import the XML data into the first worksheet starting at A1
                Workbook workbook = new Workbook();
                workbook.ImportXml(xmlStream, "Sheet1", 0, 0);

                // Retrieve the XML map that was created during import
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Get the first worksheet (where the XML was imported)
                Worksheet worksheet = workbook.Worksheets[0];

                // Query the worksheet for cells linked to a specific XML path
                // In this example we query the path of the <Item> elements
                string queryPath = "/Root/Data/Item";
                ArrayList linkedAreas = worksheet.XmlMapQuery(queryPath, xmlMap);

                // Validate the query results
                if (linkedAreas.Count == 0)
                {
                    Console.WriteLine("No cells are linked to the specified XML path.");
                }
                else
                {
                    Console.WriteLine($"Found {linkedAreas.Count} linked cell area(s) for path '{queryPath}':");
                    foreach (CellArea area in linkedAreas)
                    {
                        // For each area, output the start cell address and its current value
                        string cellName = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                        string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                        Console.WriteLine($"- Cell {cellName}: \"{cellValue}\"");
                    }
                }

                // Optionally, save the workbook to verify the import visually
                workbook.Save("XmlMapValidationResult.xlsx");
            }
        }
    }
}
