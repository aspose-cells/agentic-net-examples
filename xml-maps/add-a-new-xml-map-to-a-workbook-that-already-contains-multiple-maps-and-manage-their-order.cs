// Title: Add and reorder XML maps in an Aspose.Cells workbook using C#
// Description: Shows how to create a Workbook, add three XML maps from XSD files, preserve each map's schema URL, clear the XmlMapCollection, and re‑add the maps in a custom sequence before saving the spreadsheet.
// Keywords: Aspose.Cells | C# XML map | XmlMapCollection | add XML map | reorder XML maps | multiple XML maps | XSD schema binding | DataBinding.Url | clear XmlMapCollection | save workbook
// Common Searches: how to add an XML map to an existing Aspose.Cells workbook | reorder XmlMapCollection Aspose.Cells C# | manage multiple XML maps in a spreadsheet | change order of XML maps in Aspose.Cells | insert XML map at specific position .NET
// Developer Intent: Insert a new XML map into a workbook that already has maps and arrange all maps in a defined order.
// Use Cases: Create a workbook and import several XSD‑based XML maps for data binding. | Adjust the map order so a particular XML map appears first in the collection. | Persist the ordered maps by saving the workbook for later import or processing. | Validate the new sequence by enumerating map names before further operations.
// AI Prompts: Generate C# code that adds an XML map to an existing Aspose.Cells workbook and places it at a given index in the XmlMapCollection. | Provide a reusable method to reorder XmlMapCollection based on a list of map names without clearing the collection. | Explain how to capture each XmlMap's DataBinding URL and re‑add the maps to achieve a custom order.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapOrderDemo
{
    // Shows how to create a Workbook, add three XML maps from XSD files, preserve each map's schema URL, clear the XmlMapCollection, and re‑add the maps in a custom sequence before saving the spreadsheet.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the XmlMap collection
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Prepare temporary folder for XSD files
                string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeXmlMapDemo");
                Directory.CreateDirectory(tempFolder);

                // Helper to write schema string to a file and return the file path
                string WriteSchemaToFile(string fileName, string schemaContent)
                {
                    string filePath = Path.Combine(tempFolder, fileName);
                    File.WriteAllText(filePath, schemaContent);
                    return filePath;
                }

                // First XML map
                string schema1 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root1'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item1' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                  </xs:schema>";
                string schemaPath1 = WriteSchemaToFile("schema1.xsd", schema1);
                int index1 = xmlMaps.Add(schemaPath1);
                XmlMap map1 = xmlMaps[index1];
                map1.Name = "FirstMap";

                // Second XML map
                string schema2 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root2'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item2' type='xs:int'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                  </xs:schema>";
                string schemaPath2 = WriteSchemaToFile("schema2.xsd", schema2);
                int index2 = xmlMaps.Add(schemaPath2);
                XmlMap map2 = xmlMaps[index2];
                map2.Name = "SecondMap";

                // Third XML map (the one we want to place first in order)
                string schema3 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root3'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item3' type='xs:date'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                  </xs:schema>";
                string schemaPath3 = WriteSchemaToFile("schema3.xsd", schema3);
                int index3 = xmlMaps.Add(schemaPath3);
                XmlMap map3 = xmlMaps[index3];
                map3.Name = "ThirdMap";

                // Reorder maps: ThirdMap, FirstMap, SecondMap
                // Store current maps' schema URLs and names
                var currentMaps = new (string Url, string Name)[xmlMaps.Count];
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    currentMaps[i] = (xmlMaps[i].DataBinding.Url, xmlMaps[i].Name);
                }

                // Clear collection
                xmlMaps.Clear();

                // Add maps back in desired order
                // ThirdMap
                int idx = xmlMaps.Add(currentMaps[2].Url);
                xmlMaps[idx].Name = currentMaps[2].Name;

                // FirstMap
                idx = xmlMaps.Add(currentMaps[0].Url);
                xmlMaps[idx].Name = currentMaps[0].Name;

                // SecondMap
                idx = xmlMaps.Add(currentMaps[1].Url);
                xmlMaps[idx].Name = currentMaps[1].Name;

                // Verify order by printing map names
                Console.WriteLine("XmlMap order after reordering:");
                for (int i = 0; i < xmlMaps.Count; i++)
                {
                    Console.WriteLine($"{i}: {xmlMaps[i].Name}");
                }

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "OrderedXmlMaps.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
