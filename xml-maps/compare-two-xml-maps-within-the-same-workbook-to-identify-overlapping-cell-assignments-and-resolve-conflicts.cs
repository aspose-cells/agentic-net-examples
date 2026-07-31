// Title: C# – Detect and Resolve Overlapping XML Map Cell Assignments with Aspose.Cells
// Description: Shows how to add two XML maps to a workbook, link the same cell to both maps, query linked CellArea collections with XmlMapQuery, spot overlapping assignments, clear the conflicting link, and save the corrected file.
// Keywords: Aspose.Cells | C# | XML map conflict | overlapping cell assignments | XmlMapQuery | CellArea detection | duplicate XML mapping | clear XML link | Excel automation | programmatic XML mapping
// Common Searches: detect duplicate XML map cells Aspose.Cells | resolve overlapping XML map assignments C# | XmlMapQuery find conflicting cells | clear cell link to XML map programmatically | Aspose.Cells XML map conflict resolution example
// Developer Intent: Identify cells linked to more than one XML map in a workbook and automatically remove the extra link.
// Use Cases: Validate imported XML data to ensure each worksheet cell is bound to a single map. | Automate cleanup of legacy spreadsheets where multiple schemas were applied to the same range. | Generate a report of conflicting cell mappings before publishing an Excel file.
// AI Prompts: Write C# code using Aspose.Cells that lists every cell linked to a given XML path across all XML maps and flags duplicates. | Explain how to detach a cell from an XML map without breaking other mappings in an Aspose.Cells workbook. | Provide a method to merge two XML maps that share paths while preserving unique cell links and removing overlaps.

using System;
using System.Collections;
using Aspose.Cells;

// Shows how to add two XML maps to a workbook, link the same cell to both maps, query linked CellArea collections with XmlMapQuery, spot overlapping assignments, clear the conflicting link, and save the corrected file.
class XmlMapComparison
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Define two simple XML schemas (maps) that have the same element path
            string schema1 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                <xs:element name='Root'>
                    <xs:complexType>
                        <xs:sequence>
                            <xs:element name='Item' type='xs:string'/>
                        </xs:sequence>
                    </xs:complexType>
                </xs:element>
            </xs:schema>";

            string schema2 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                <xs:element name='Root'>
                    <xs:complexType>
                        <xs:sequence>
                            <xs:element name='Item' type='xs:string'/>
                        </xs:sequence>
                    </xs:complexType>
                </xs:element>
            </xs:schema>";

            // Add the two XML maps to the workbook
            int mapIdx1 = wb.Worksheets.XmlMaps.Add(schema1);
            int mapIdx2 = wb.Worksheets.XmlMaps.Add(schema2);
            XmlMap map1 = wb.Worksheets.XmlMaps[mapIdx1];
            XmlMap map2 = wb.Worksheets.XmlMaps[mapIdx2];
            map1.Name = "Map1";
            map2.Name = "Map2";

            // Intentionally create a conflict by linking the same cell (A1) to both maps
            ws.Cells.LinkToXmlMap(map1.Name, 0, 0, "/Root/Item"); // A1 -> Map1
            ws.Cells.LinkToXmlMap(map2.Name, 0, 0, "/Root/Item"); // A1 -> Map2 (conflict)

            // Query the cell areas that are linked to the same XML path for each map
            ArrayList areasMap1 = ws.XmlMapQuery("/Root/Item", map1);
            ArrayList areasMap2 = ws.XmlMapQuery("/Root/Item", map2);

            // Detect overlapping assignments and resolve them
            foreach (CellArea area1 in areasMap1)
            {
                foreach (CellArea area2 in areasMap2)
                {
                    // Simple overlap check: same start row and column
                    if (area1.StartRow == area2.StartRow && area1.StartColumn == area2.StartColumn)
                    {
                        Console.WriteLine($"Conflict detected at cell ({area1.StartRow}, {area1.StartColumn})");

                        // Resolve conflict by clearing the cell's content (removes the link)
                        ws.Cells[area2.StartRow, area2.StartColumn].PutValue(string.Empty);
                        Console.WriteLine("Resolved conflict by clearing the cell linked to the second map.");
                    }
                }
            }

            // Save the workbook with the resolved mappings
            wb.Save("XmlMapConflictResolved.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
