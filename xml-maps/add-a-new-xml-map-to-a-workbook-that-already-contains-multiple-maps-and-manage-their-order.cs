using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapOrderDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -----------------------------------------------------------------
                // Prepare temporary XSD files for the XML maps
                // -----------------------------------------------------------------
                string tempDir = Path.Combine(Path.GetTempPath(), "AsposeXmlMaps");
                Directory.CreateDirectory(tempDir);

                string schema1Path = Path.Combine(tempDir, "schema1.xsd");
                string schema2Path = Path.Combine(tempDir, "schema2.xsd");
                string schemaNewPath = Path.Combine(tempDir, "schemaNew.xsd");

                string schema1 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root1'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item1' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                  </xs:schema>";
                string schema2 = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root2'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item2' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                  </xs:schema>";
                string schemaNew = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                     <xs:element name='RootNew'>
                                         <xs:complexType>
                                             <xs:sequence>
                                                 <xs:element name='ItemNew' type='xs:string'/>
                                             </xs:sequence>
                                         </xs:complexType>
                                     </xs:element>
                                   </xs:schema>";

                File.WriteAllText(schema1Path, schema1);
                File.WriteAllText(schema2Path, schema2);
                File.WriteAllText(schemaNewPath, schemaNew);

                // -----------------------------------------------------------------
                // Add initial XML maps (simulating a workbook that already has maps)
                // -----------------------------------------------------------------
                if (File.Exists(schema1Path))
                {
                    int index1 = workbook.Worksheets.XmlMaps.Add(schema1Path);
                    workbook.Worksheets.XmlMaps[index1].Name = "FirstMap";
                }

                if (File.Exists(schema2Path))
                {
                    int index2 = workbook.Worksheets.XmlMaps.Add(schema2Path);
                    workbook.Worksheets.XmlMaps[index2].Name = "SecondMap";
                }

                // ---------------------------------------------------------------
                // Add a new XML map that should appear between the existing maps
                // ---------------------------------------------------------------
                int newIndex = -1;
                if (File.Exists(schemaNewPath))
                {
                    newIndex = workbook.Worksheets.XmlMaps.Add(schemaNewPath);
                    workbook.Worksheets.XmlMaps[newIndex].Name = "InsertedMap";
                }

                // ---------------------------------------------------------------
                // Reorder: move the newly added map to position 1 (second place)
                // ---------------------------------------------------------------
                if (newIndex >= 0)
                {
                    // Remove the map from its current position
                    workbook.Worksheets.XmlMaps.RemoveAt(newIndex);

                    // Store existing maps temporarily
                    XmlMap[] existingMaps = new XmlMap[workbook.Worksheets.XmlMaps.Count];
                    for (int i = 0; i < existingMaps.Length; i++)
                        existingMaps[i] = workbook.Worksheets.XmlMaps[i];

                    // Clear the collection
                    workbook.Worksheets.XmlMaps.Clear();

                    // Re‑add maps in the required order:
                    // 1. FirstMap
                    // 2. InsertedMap (new map)
                    // 3. SecondMap
                    if (File.Exists(schema1Path))
                        workbook.Worksheets.XmlMaps.Add(schema1Path);
                    if (File.Exists(schemaNewPath))
                        workbook.Worksheets.XmlMaps.Add(schemaNewPath);
                    if (File.Exists(schema2Path))
                        workbook.Worksheets.XmlMaps.Add(schema2Path);

                    // Restore map names (names are lost after Clear/Add)
                    if (workbook.Worksheets.XmlMaps.Count > 0)
                        workbook.Worksheets.XmlMaps[0].Name = "FirstMap";
                    if (workbook.Worksheets.XmlMaps.Count > 1)
                        workbook.Worksheets.XmlMaps[1].Name = "InsertedMap";
                    if (workbook.Worksheets.XmlMaps.Count > 2)
                        workbook.Worksheets.XmlMaps[2].Name = "SecondMap";
                }

                // ---------------------------------------------------------------
                // Display the current order of XML maps
                // ---------------------------------------------------------------
                Console.WriteLine("Current XML map order:");
                for (int i = 0; i < workbook.Worksheets.XmlMaps.Count; i++)
                {
                    Console.WriteLine($"{i}: {workbook.Worksheets.XmlMaps[i].Name}");
                }

                // Save the workbook (the XML maps are stored inside the file)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithOrderedXmlMaps.xlsx");
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