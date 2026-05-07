using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (supported format such as .xlsx)
            Workbook workbook = new Workbook("input.xlsx");

            // Define an XML schema (XSD) as a string.
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file because Aspose.Cells expects a file path.
            string schemaPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
            File.WriteAllText(schemaPath, xmlSchema);

            // Add the XML map to the workbook via the XmlMaps collection.
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);

            // Optionally, set a friendly name for the map.
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SampleMap";

            // Save the workbook with the newly added XML map.
            workbook.Save("output.xlsx");

            // Clean up the temporary schema file.
            if (File.Exists(schemaPath))
                File.Delete(schemaPath);
        }
    }
}