using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -----------------------------------------------------------------
                // 1. Define an XML schema (XSD) that describes the XML structure.
                //    This schema will be used to create an XML map in the workbook.
                // -----------------------------------------------------------------
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Transmittals'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Issued_Document' type='xs:string'/>
                                                    <xs:element name='Date' type='xs:date'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Write the schema to a temporary file because Aspose.Cells expects a file path
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Add the XML map to the workbook and give it a friendly name
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "Transmittals_Map";

                // Clean up the temporary schema file
                if (File.Exists(tempSchemaPath))
                {
                    File.Delete(tempSchemaPath);
                }

                // -----------------------------------------------------------------
                // 2. Import an initial XML file so that the workbook contains data.
                //    The XML file can be created on the fly for this demo.
                // -----------------------------------------------------------------
                string initialXml = @"<Transmittals>
                                        <Issued_Document>InitialDoc</Issued_Document>
                                        <Date>2023-01-01</Date>
                                     </Transmittals>";

                using (MemoryStream xmlStream = new MemoryStream())
                using (StreamWriter writer = new StreamWriter(xmlStream))
                {
                    writer.Write(initialXml);
                    writer.Flush();
                    xmlStream.Position = 0;

                    // Import the XML data starting at cell A1 of the first worksheet
                    workbook.ImportXml(xmlStream, "Sheet1", 0, 0);
                }

                // -----------------------------------------------------------------
                // 3. Link specific cells to elements of the XML map.
                //    Changes made to these cells will be reflected in the XML.
                // -----------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Link cell A2 (row 1, column 0) to the Issued_Document element
                cells.LinkToXmlMap("Transmittals_Map", 1, 0, "/Transmittals/Issued_Document");

                // Link cell A3 (row 2, column 0) to the Date element
                cells.LinkToXmlMap("Transmittals_Map", 2, 0, "/Transmittals/Date");

                // -----------------------------------------------------------------
                // 4. Update the linked cells with new values.
                //    The underlying XML map will capture these changes.
                // -----------------------------------------------------------------
                cells["A2"].PutValue("UpdatedDoc");
                cells["A3"].PutValue(new DateTime(2024, 12, 31));

                // -----------------------------------------------------------------
                // 5. Export the updated XML to verify that changes are reflected.
                // -----------------------------------------------------------------
                string outputXmlPath = "UpdatedTransmittals.xml";
                workbook.ExportXml(xmlMap.Name, outputXmlPath);

                // -----------------------------------------------------------------
                // 6. Save the workbook (lifecycle rule: use provided save method only)
                // -----------------------------------------------------------------
                string workbookPath = "MappedWorkbook.xlsx";
                workbook.Save(workbookPath);

                // Inform the user
                Console.WriteLine($"Workbook saved as '{workbookPath}'.");
                Console.WriteLine($"Updated XML exported to '{outputXmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}