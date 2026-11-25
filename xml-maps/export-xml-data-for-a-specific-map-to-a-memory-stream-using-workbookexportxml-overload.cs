// Title: Export an XML map to a MemoryStream using Workbook.ExportXml in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells export XML map to MemoryStream C# example | Workbook.ExportXml overload that writes XML to a stream in .NET | How to retrieve XML map data without saving a workbook using Aspose.Cells | C# code to serialize an XML map directly into a MemoryStream with Aspose.Cells | In‑memory XML export from Excel workbook using Aspose.Cells API
// Tags: Workbook.ExportXml to MemoryStream | XML map serialization Aspose.Cells C# | in‑memory XML export without file | Aspose.Cells XML map stream output | C# memory stream XML generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlToMemoryStreamDemo
{
    // The sample creates a workbook, adds sample data, defines an XML schema, registers it as an XML map, and then uses Workbook.ExportXml with the map name to write the XML representation directly into a MemoryStream. The stream is read back for display and optionally saved to a file, demonstrating how to export XML map data in‑memory without persisting the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for a temporary workbook file (will be deleted at the end)
            string tempWorkbookPath = Path.Combine(Path.GetTempPath(), "DemoWorkbook.xlsx");

            // Simple XML schema (XSD) that defines a root element <Data> with repeating <Item> elements
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Data'>
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name='Item' maxOccurs='unbounded'>
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name='Id' type='xs:int'/>
                              <xs:element name='Name' type='xs:string'/>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and add some sample data
                // ------------------------------------------------------------
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "DataSheet";

                    // Header row
                    sheet.Cells["A1"].PutValue("Id");
                    sheet.Cells["B1"].PutValue("Name");

                    // Sample rows
                    sheet.Cells["A2"].PutValue(1);
                    sheet.Cells["B2"].PutValue("Alice");

                    sheet.Cells["A3"].PutValue(2);
                    sheet.Cells["B3"].PutValue("Bob");

                    // ------------------------------------------------------------
                    // 2. Add the XML map to the workbook using the schema defined above
                    // ------------------------------------------------------------
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                    // For demonstration we simply use the map name; binding cells to the map
                    // would normally be done via ListObject or ImportXml, but ExportXml works
                    // even if no explicit bindings exist (it will produce an empty XML document
                    // based on the map definition).
                    string mapName = xmlMap.Name;
                    Console.WriteLine($"Created XML map with name: {mapName}");

                    // ------------------------------------------------------------
                    // 3. Save the workbook to a temporary file (optional, just to show it exists)
                    // ------------------------------------------------------------
                    workbook.Save(tempWorkbookPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to temporary file: {tempWorkbookPath}");

                    // ------------------------------------------------------------
                    // 4. Export the XML data linked by the map to a MemoryStream
                    // ------------------------------------------------------------
                    using (MemoryStream xmlStream = new MemoryStream())
                    {
                        // ExportXml overload that writes to a Stream
                        workbook.ExportXml(mapName, xmlStream);

                        // Reset stream position to read from the beginning
                        xmlStream.Position = 0;

                        // Read the exported XML as a string for display
                        using (StreamReader reader = new StreamReader(xmlStream))
                        {
                            string exportedXml = reader.ReadToEnd();
                            Console.WriteLine("\n--- Exported XML Content ---");
                            Console.WriteLine(exportedXml);
                            Console.WriteLine("--- End of Exported XML ---\n");
                        }

                        // (Optional) Write the same XML to a physical file for verification
                        string outputXmlPath = Path.Combine(Path.GetDirectoryName(tempWorkbookPath), "ExportedData.xml");
                        File.WriteAllBytes(outputXmlPath, xmlStream.ToArray());
                        Console.WriteLine($"Exported XML also written to file: {outputXmlPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                // Clean up the temporary workbook file if it exists
                if (File.Exists(tempWorkbookPath))
                {
                    try
                    {
                        File.Delete(tempWorkbookPath);
                        Console.WriteLine($"Temporary workbook file deleted: {tempWorkbookPath}");
                    }
                    catch (Exception cleanupEx)
                    {
                        Console.WriteLine($"Failed to delete temporary file: {cleanupEx.Message}");
                    }
                }
            }
        }
    }
}
