using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidateXmlMappingBeforeExport
    {
        public static void Run()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Define a simple XML schema with required elements
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Id' type='xs:int'/>
                                                    <xs:element name='Name' type='xs:string'/>
                                                    <xs:element name='Email' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // 3. Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "UserDataMap";

                // 4. Get the first worksheet and its cells
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 5. Link worksheet cells to XML elements (Id and Name are mapped, Email is intentionally left unmapped)
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Id");      // Cell A1 -> Id
                cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Root/Name");    // Cell B1 -> Name
                // Email element is not linked to any cell to demonstrate validation failure

                // 6. Populate the linked cells with sample data
                cells["A1"].PutValue(101);
                cells["B1"].PutValue("Alice");

                // 7. Define the list of required XML element paths that must have a mapped cell
                string[] requiredPaths = new string[]
                {
                    "/Root/Id",
                    "/Root/Name",
                    "/Root/Email"
                };

                // 8. Validate that each required path has at least one mapped cell
                List<string> missingMappings = new List<string>();
                foreach (string path in requiredPaths)
                {
                    // Query the worksheet for cell areas linked to the current path
                    ArrayList cellAreas = sheet.XmlMapQuery(path, xmlMap);
                    if (cellAreas == null || cellAreas.Count == 0)
                    {
                        missingMappings.Add(path);
                    }
                }

                // 9. If any required element is missing a mapping, report and abort export
                if (missingMappings.Count > 0)
                {
                    Console.WriteLine("Validation failed. The following XML elements have no mapped cells:");
                    foreach (string missing in missingMappings)
                    {
                        Console.WriteLine($" - {missing}");
                    }
                    Console.WriteLine("Export aborted.");
                    return;
                }

                // 10. All required elements are mapped; proceed to export the XML data
                try
                {
                    workbook.ExportXml(xmlMap.Name, "ExportedData.xml");
                    Console.WriteLine("XML exported successfully to ExportedData.xml");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during ExportXml: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateXmlMappingBeforeExport.Run();
        }
    }
}