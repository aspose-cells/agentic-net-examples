using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapSummary
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // ------------------------------------------------------------
                // Sample data: add two XML maps and link some cells to them
                // ------------------------------------------------------------
                // First XML map (simple schema)
                string xmlSchema1 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Write schema to a temporary file because XmlMaps.Add expects a file path
                string tempFile1 = Path.GetTempFileName();
                File.WriteAllText(tempFile1, xmlSchema1);
                int mapIndex1 = workbook.Worksheets.XmlMaps.Add(tempFile1);
                XmlMap map1 = workbook.Worksheets.XmlMaps[mapIndex1];
                map1.Name = "FirstMap";

                // Link a cell to the first map
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Cells["A1"].PutValue("Item");
                sheet1.Cells.LinkToXmlMap(map1.Name, 0, 0, "/Root/Item");

                // Second XML map (different schema)
                string xmlSchema2 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Data'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Value' type='xs:int'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                string tempFile2 = Path.GetTempFileName();
                File.WriteAllText(tempFile2, xmlSchema2);
                int mapIndex2 = workbook.Worksheets.XmlMaps.Add(tempFile2);
                XmlMap map2 = workbook.Worksheets.XmlMaps[mapIndex2];
                map2.Name = "SecondMap";

                // Add a second worksheet and link a cell to the second map
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                sheet2.Cells["B2"].PutValue(123);
                sheet2.Cells.LinkToXmlMap(map2.Name, 1, 1, "/Data/Value");

                // ------------------------------------------------------------
                // Create a summary worksheet
                // ------------------------------------------------------------
                Worksheet summarySheet = workbook.Worksheets.Add("XmlMapSummary");
                // Write header
                summarySheet.Cells["A1"].PutValue("Map Name");
                summarySheet.Cells["B1"].PutValue("Root Element");
                summarySheet.Cells["C1"].PutValue("Linked Cells Count");

                int summaryRow = 1; // zero‑based index (row 2 in Excel)

                // ------------------------------------------------------------
                // Iterate through all XML maps and compute linked cells
                // ------------------------------------------------------------
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    XmlMap map = xmlMaps[i];
                    string mapName = map.Name ?? $"Map_{i}";
                    string rootElement = map.RootElementName ?? string.Empty;

                    int linkedCellsCount = 0;

                    // Query each worksheet (except the summary sheet) for cells linked to this map
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        if (ws.Name == summarySheet.Name) continue; // skip summary sheet

                        // Use the root element as the query path; prepend '/' to form a valid XPath
                        string queryPath = string.IsNullOrEmpty(rootElement) ? string.Empty : "/" + rootElement;

                        // XmlMapQuery returns an ArrayList of CellArea objects
                        ArrayList cellAreas = ws.XmlMapQuery(queryPath, map);
                        if (cellAreas != null)
                        {
                            linkedCellsCount += cellAreas.Count;
                        }
                    }

                    // Write the information into the summary sheet
                    summarySheet.Cells[summaryRow, 0].PutValue(mapName);
                    summarySheet.Cells[summaryRow, 1].PutValue(rootElement);
                    summarySheet.Cells[summaryRow, 2].PutValue(linkedCellsCount);
                    summaryRow++;
                }

                // ------------------------------------------------------------
                // Save the workbook (lifecycle: save)
                // ------------------------------------------------------------
                workbook.Save("XmlMapSummaryReport.xlsx");
                // Clean up temporary schema files
                if (File.Exists(tempFile1)) File.Delete(tempFile1);
                if (File.Exists(tempFile2)) File.Delete(tempFile2);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}