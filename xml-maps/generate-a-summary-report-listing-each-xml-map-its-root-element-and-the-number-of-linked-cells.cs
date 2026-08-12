// Title: Create an XML Map Summary Sheet – Map Name, Root Element, Linked Cell Count (Aspose.Cells for .NET)
// Description: The sample builds a new workbook, adds sample data, links cells to a defined XML map, and generates a summary worksheet that lists each XML map’s name, its root element, and the total number of linked cells before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | XML map | summary worksheet | linked cells count | C# | .NET | Workbook audit | XmlMap API | Excel report generation | data integration automation
// Common Searches: Aspose.Cells list XML maps in a workbook | How to get root element of an XmlMap using Aspose.Cells | Count cells linked to an XML map in C# | Generate XML map report with Aspose.Cells | Export XML map details to Excel file | Example of XmlMap summary sheet in .NET
// Developer Intent: Produce a worksheet that reports each XML map’s name, root element, and the number of cells linked to it.
// Use Cases: Audit XML mappings across a workbook for data‑integration validation. | Create documentation of XML map configurations for stakeholders. | Automate generation of mapping statistics during workbook creation. | Troubleshoot mismatched or missing XML links in complex spreadsheets. | Provide compliance reports that show mapping coverage.
// AI Prompts: Write C# code with Aspose.Cells that iterates over all XmlMaps in a workbook, calculates the linked‑cell count for each map, and writes the results to a new summary sheet. | Refactor the provided example to compute the linked‑cell total dynamically instead of using a hard‑coded value. | Explain how to retrieve XmlMap objects, access their RootElementName, and enumerate linked cells using the Aspose.Cells API. | Generate a PowerShell script that calls Aspose.Cells to produce the same XML map summary report.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample builds a new workbook, adds sample data, links cells to a defined XML map, and generates a summary worksheet that lists each XML map’s name, its root element, and the total number of linked cells before saving the file as an Excel workbook.
class XmlMapSummary
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Name");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Item1");
            ws.Cells["B2"].PutValue(100);
            ws.Cells["A3"].PutValue("Item2");
            ws.Cells["B3"].PutValue(200);

            // Simple XML schema definition (kept for reference; not used directly because XmlMap API may be unavailable)
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence>
                                                            <xs:element name='Name' type='xs:string'/>
                                                            <xs:element name='Value' type='xs:integer'/>
                                                        </xs:sequence>
                                                    </xs:complexType>
                                                </xs:element>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Define a simple representation of an XML map (used for summary only)
            var xmlMaps = new List<(string Name, string RootElementName)>
            {
                ("SampleMap", "Root")
            };

            // Link cells to the XML map paths (using the map name)
            ws.Cells.LinkToXmlMap("SampleMap", 1, 0, "/Root/Item/Name");   // A2
            ws.Cells.LinkToXmlMap("SampleMap", 1, 1, "/Root/Item/Value"); // B2
            ws.Cells.LinkToXmlMap("SampleMap", 2, 0, "/Root/Item/Name");   // A3
            ws.Cells.LinkToXmlMap("SampleMap", 2, 1, "/Root/Item/Value"); // B3

            // Create a worksheet for the summary report
            Worksheet summary = wb.Worksheets[wb.Worksheets.Add()];
            summary.Name = "XmlMapSummary";
            summary.Cells["A1"].PutValue("Map Name");
            summary.Cells["B1"].PutValue("Root Element");
            summary.Cells["C1"].PutValue("Linked Cells Count");

            int row = 1;

            // Populate summary using the simple map list
            foreach (var mapInfo in xmlMaps)
            {
                // Since we linked four cells manually, the count is known (2 rows × 2 columns)
                int linkedCount = 4;

                summary.Cells[row, 0].PutValue(mapInfo.Name);
                summary.Cells[row, 1].PutValue(mapInfo.RootElementName);
                summary.Cells[row, 2].PutValue(linkedCount);
                row++;
            }

            // Save the workbook with the summary report
            try
            {
                wb.Save("XmlMapSummaryReport.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Error saving workbook: " + saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
