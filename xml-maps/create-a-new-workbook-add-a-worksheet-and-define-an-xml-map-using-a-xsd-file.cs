using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet to the workbook
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet worksheet = workbook.Worksheets[sheetIndex];
            worksheet.Name = "DataSheet";

            // Define an XSD schema as a string
            string xsdSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                 </xs:schema>";

            // Write the schema to a temporary file (required by Aspose.Cells API)
            string tempXsdPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
            File.WriteAllText(tempXsdPath, xsdSchema);

            // Ensure the temporary XSD file exists before adding the XML map
            if (File.Exists(tempXsdPath))
            {
                // Add the XML map to the workbook using the schema file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "RootMap";
            }
            else
            {
                throw new FileNotFoundException("Temporary XSD schema file was not created.", tempXsdPath);
            }

            // Save the workbook
            string outputPath = "WorkbookWithXmlMap.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}