// Title: Find a Cell Mapped to an XML Attribute Using Worksheet.Cells.Find (Aspose.Cells .NET)
// Description: Demonstrates how to import XML into a workbook, create an XML map, link a worksheet cell to an attribute, search for the attribute value with Worksheet.Cells.Find, verify the mapping via XmlMapQuery, and save the result.
// Keywords: Aspose.Cells | Worksheet.Cells.Find | XML map | LinkToXmlMap | XmlMapQuery | C# XML attribute lookup | search mapped cell | Aspose.Cells .NET example
// Common Searches: Aspose.Cells find cell by XML attribute | Worksheet.Cells.Find XML map C# | link cell to XML attribute Aspose | verify XML mapped cell Aspose.Cells | search XML attribute value in workbook
// Developer Intent: Locate the worksheet cell that is linked to a specific XML attribute and confirm that the found cell belongs to the XML map.
// Use Cases: Confirm that a cell linked to an XML attribute holds the expected value before further processing. | Programmatically locate and update cells representing XML attribute values during data synchronization. | Generate a list of cell addresses for all occurrences of a particular XML attribute after a search operation.
// AI Prompts: Write C# code that imports XML into a worksheet, links a cell to an attribute, searches for the attribute value using Worksheet.Cells.Find, and validates the mapping with XmlMapQuery. | Explain step‑by‑step how Worksheet.Cells.Find and XmlMapQuery can be combined to ensure a found cell is part of a specific XML map in Aspose.Cells for .NET.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsXmlFindDemo
{
    // Demonstrates how to import XML into a workbook, create an XML map, link a worksheet cell to an attribute, search for the attribute value with Worksheet.Cells.Find, verify the mapping via XmlMapQuery, and save the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Define a simple XML with an attribute we want to map
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Item id='123'>SampleValue</Item>
</Root>";

                // 3. Import the XML into the first worksheet.
                //    This creates an XML map and populates the sheet with the XML data.
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // 4. Retrieve the created XML map.
                //    In newer Aspose.Cells versions the collection is accessed via workbook.XmlMaps.
                //    For compatibility, use the XmlMapCollection property if XmlMaps is unavailable.
                XmlMap xmlMap = null;
                // Try the standard XmlMaps property first
                var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                if (xmlMapsProp != null)
                {
                    var xmlMaps = (XmlMapCollection)xmlMapsProp.GetValue(workbook);
                    if (xmlMaps.Count > 0)
                        xmlMap = xmlMaps[0];
                }
                else
                {
                    // Fallback to XmlMapCollection (older API)
                    var xmlMapCollProp = workbook.GetType().GetProperty("XmlMapCollection");
                    if (xmlMapCollProp != null)
                    {
                        var xmlMaps = (XmlMapCollection)xmlMapCollProp.GetValue(workbook);
                        if (xmlMaps.Count > 0)
                            xmlMap = xmlMaps[0];
                    }
                }

                if (xmlMap == null)
                {
                    Console.WriteLine("No XML map was created.");
                    return;
                }

                // 5. Link a specific cell to the XML attribute "id"
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;
                cells.LinkToXmlMap(xmlMap.Name, 2, 0, "/Root/Item/@id"); // Links cell A3 to the attribute

                // 6. Ensure the linked cell contains the expected value.
                cells["A3"].PutValue("123");

                // 7. Use Find to locate the cell that contains the attribute value "123"
                Cell foundCell = sheet.Cells.Find("123", null);

                // 8. Verify that the found cell is part of the XML mapping
                ArrayList mappedAreas = sheet.XmlMapQuery("/Root/Item/@id", xmlMap);
                bool isMapped = false;
                if (foundCell != null)
                {
                    foreach (CellArea area in mappedAreas)
                    {
                        if (foundCell.Row >= area.StartRow && foundCell.Row <= area.EndRow &&
                            foundCell.Column >= area.StartColumn && foundCell.Column <= area.EndColumn)
                        {
                            isMapped = true;
                            break;
                        }
                    }
                }

                // 9. Output the result
                if (foundCell != null && isMapped)
                {
                    Console.WriteLine($"Found mapped cell at {foundCell.Name} with value '{foundCell.StringValue}'.");
                }
                else
                {
                    Console.WriteLine("Mapped cell not found.");
                }

                // 10. Save the workbook
                workbook.Save("XmlAttributeFindDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
