// Title: Validate Required XML Elements Are Mapped to Cells Before Exporting with Aspose.Cells for .NET
// Description: This C# example shows how to add an XML map from a temporary XSD, link worksheet cells to required elements, and use XmlMapQuery to verify that every mandatory XML node (Id, Name, Email) has at least one mapped cell. If any required mapping is missing, the export is cancelled and a clear message is logged.
// Keywords: Aspose.Cells | XML map validation | C# .NET | XmlMapQuery | required XML elements | cell-to-XML mapping | ExportXml | XSD schema | missing mapping detection | GitHub Aspose.Cells example
// Common Searches: how to check required XML elements are linked to cells using Aspose.Cells | Aspose.Cells .NET validate XML map before export | XmlMapQuery missing element detection C# | export XML only when all required mappings exist Aspose | C# example for XML map validation with Aspose.Cells | GitHub Aspose.Cells XML map validation sample
// Developer Intent: Confirm that every element defined as required in the XML schema has at least one worksheet cell linked before calling ExportXml.
// Use Cases: Prevent generation of invalid XML by aborting export when mandatory fields lack a cell mapping. | Provide developers with a validation routine that logs unmapped required elements for quick correction. | Enable automated quality checks in data‑export pipelines that rely on Aspose.Cells XML maps.
// AI Prompts: Write a C# method for Aspose.Cells that receives a Workbook, an XmlMap name, and a list of required XPath strings, then returns a list of paths that have no linked cells. | Generate a refactored version of the validation loop that throws a custom MissingXmlMappingException containing all unmapped element names. | Create a reusable utility class in .NET that validates XML map completeness and integrates with Aspose.Cells ExportXml, including logging and optional workbook saving.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlValidationDemo
{
    // This C# example shows how to add an XML map from a temporary XSD, link worksheet cells to required elements, and use XmlMapQuery to verify that every mandatory XML node (Id, Name, Email) has at least one mapped cell. If any required mapping is missing, the export is cancelled and a clear message is logged.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a simple XML schema with required elements
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

                // Write the schema to a temporary file (required by Aspose.Cells API)
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Ensure the temporary schema file exists before adding the XML map
                if (!File.Exists(tempSchemaPath))
                {
                    Console.WriteLine("Failed to create temporary schema file.");
                    return;
                }

                // Add the XML map to the workbook using the temporary schema file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "UserDataMap";

                // Get the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Link worksheet cells to XML elements
                // A1 -> /Root/Id
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Id");
                // B1 -> /Root/Name
                cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Root/Name");
                // C1 -> /Root/Email   (intentionally omitted to demonstrate validation)
                // cells.LinkToXmlMap(xmlMap.Name, 0, 2, "/Root/Email");

                // Populate some sample data
                cells[0, 0].PutValue(101);      // Id
                cells[0, 1].PutValue("Alice"); // Name
                // Email cell left unmapped on purpose

                // List of required XML element paths that must have a mapped cell
                string[] requiredPaths = new string[]
                {
                    "/Root/Id",
                    "/Root/Name",
                    "/Root/Email"
                };

                // Validate that each required path has at least one mapped cell
                bool allMapped = true;
                foreach (string path in requiredPaths)
                {
                    // Query the worksheet for cells linked to the current XML path
                    ArrayList mappedAreas = sheet.XmlMapQuery(path, xmlMap);

                    if (mappedAreas.Count == 0)
                    {
                        Console.WriteLine($"Validation failed: No cell is mapped to required XML element '{path}'.");
                        allMapped = false;
                    }
                }

                // Export XML only if validation succeeds
                if (allMapped)
                {
                    // Export the XML data using the map name
                    workbook.ExportXml(xmlMap.Name, "ExportedData.xml");
                    Console.WriteLine("XML exported successfully to 'ExportedData.xml'.");
                }
                else
                {
                    Console.WriteLine("XML export aborted due to missing mappings.");
                }

                // Optionally save the workbook for inspection
                workbook.Save("WorkbookWithMappings.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
