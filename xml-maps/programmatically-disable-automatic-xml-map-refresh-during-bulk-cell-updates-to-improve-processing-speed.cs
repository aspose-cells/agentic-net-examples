// Title: How to disable automatic XML map refresh in Aspose.Cells for .NET during high‑volume cell writes
// Common Searches: Aspose.Cells disable XML map refresh when writing thousands of rows in C# | C# bulk cell update performance without triggering XML map refresh Aspose.Cells | turn off automatic XML map refresh in Aspose.Cells .NET API | optimize Excel export speed by preventing XML map refresh in Aspose.Cells | Aspose.Cells workbook ExportXml without refreshing XML map
// Tags: suppress xml map refresh Aspose.Cells | high‑speed bulk cell writes Aspose.Cells | xml export performance Aspose.Cells | Aspose.Cells XML map optimization techniques | large Excel dataset handling Aspose.Cells

using System;
using Aspose.Cells;

namespace DisableXmlMapRefreshDemo
{
    // This C# example demonstrates adding an XML map to a workbook, writing 10,000 rows of data without invoking a map refresh, and then exporting the data to XML, illustrating how to avoid the overhead of automatic XML‑map refresh during bulk cell updates to improve performance.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the output workbook
                const string outputPath = "BulkUpdate_NoXmlMapRefresh.xlsx";

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 1. Create a simple XML schema and add it as an XML map to the workbook
                // -----------------------------------------------------------------
                string xmlSchema = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                      <xs:element name='Data'>
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

                // Add the XML map to the workbook's XmlMaps collection
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "DemoMap";

                // -----------------------------------------------------------------
                // 2. Perform bulk cell updates (e.g., write 10,000 rows)
                // -----------------------------------------------------------------
                const int rowCount = 10000;
                Cells cells = sheet.Cells;

                // Write header
                cells["A1"].PutValue("Name");
                cells["B1"].PutValue("Value");

                // Bulk insert data
                for (int i = 0; i < rowCount; i++)
                {
                    int row = i + 2; // start from row 2 (after header)
                    cells[$"A{row}"].PutValue($"Item_{i + 1}");
                    cells[$"B{row}"].PutValue(i * 10);
                }

                // -----------------------------------------------------------------
                // 3. Export data to XML using the map (refreshes the map automatically)
                // -----------------------------------------------------------------
                // The ExportXml method internally refreshes the XML map, so an explicit
                // RefreshXmlMap call is not required.
                workbook.ExportXml(xmlMap.Name, "ExportedData.xml");

                // -----------------------------------------------------------------
                // 4. Save the workbook
                // -----------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("Bulk updates completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
