using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlImportDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Define a simple XML schema (XSD) as a string.
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Products'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Product' maxOccurs='unbounded'>
                                                        <xs:complexType>
                                                            <xs:sequence>
                                                                <xs:element name='Name' type='xs:string'/>
                                                                <xs:element name='Price' type='xs:decimal'/>
                                                            </xs:sequence>
                                                        </xs:complexType>
                                                    </xs:element>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // Write the schema to a temporary file because Aspose.Cells expects a file path.
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Ensure the temporary schema file exists before adding the XML map.
                if (!File.Exists(tempSchemaPath))
                    throw new FileNotFoundException("Temporary XSD file was not created.", tempSchemaPath);

                // Add the XML map to the workbook using the schema file.
                int mapIndex = wb.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "ProductsMap";

                // Get the first worksheet and its cells collection
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Link cells to the XML map paths
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Products/Product/Name");
                cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Products/Product/Price");

                // Path to the XML data file
                string xmlDataPath = "data.xml";

                // Verify that the XML data file exists before importing
                if (!File.Exists(xmlDataPath))
                    throw new FileNotFoundException("XML data file not found.", xmlDataPath);

                // Import XML data into the worksheet starting at cell A1.
                wb.ImportXml(xmlDataPath, sheet.Name, 0, 0);

                // Save the workbook with the linked cells populated from the XML data
                string outputPath = "ProductsOutput.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}