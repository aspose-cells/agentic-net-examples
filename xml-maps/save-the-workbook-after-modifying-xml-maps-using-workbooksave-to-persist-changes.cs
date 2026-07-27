using System;
using System.IO;
using Aspose.Cells;

class SaveWorkbookWithXmlMap
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // Define a simple XML schema that will be used for the map
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file (required by Aspose.Cells API)
            string schemaPath = "SampleSchema.xsd";
            if (!File.Exists(schemaPath))
            {
                File.WriteAllText(schemaPath, xmlSchema);
            }

            // Add the XML map to the workbook's XmlMaps collection using the schema file path
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Modify the XML map (e.g., change its name)
            xmlMap.Name = "MySampleMap";

            // Create XmlSaveOptions and specify the map to be exported when saving
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                XmlMapName = xmlMap.Name
            };

            // Save the workbook; the modified XML map is persisted in the file
            string outputPath = "WorkbookWithXmlMap.xml";
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}