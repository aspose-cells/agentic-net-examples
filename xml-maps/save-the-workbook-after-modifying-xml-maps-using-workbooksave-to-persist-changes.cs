using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook instance (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet to avoid possible name conflicts with XML maps
            Worksheet defaultSheet = workbook.Worksheets[0];
            defaultSheet.Name = "DataSheet";

            // Simple XML schema used for the map
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file
            string tempSchemaPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
            File.WriteAllText(tempSchemaPath, xmlSchema);

            // Verify the schema file exists before adding the XML map
            if (!File.Exists(tempSchemaPath))
                throw new FileNotFoundException("XML schema file not found.", tempSchemaPath);

            // Add the XML map to the workbook's XmlMaps collection
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Ensure the map name does not clash with existing worksheet names
            string desiredMapName = "MySampleMap";
            bool nameExists = false;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(desiredMapName, StringComparison.OrdinalIgnoreCase))
                {
                    nameExists = true;
                    break;
                }
            }
            xmlMap.Name = nameExists ? desiredMapName + "_Map" : desiredMapName;

            // Save the workbook
            string outputPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithXmlMap.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}