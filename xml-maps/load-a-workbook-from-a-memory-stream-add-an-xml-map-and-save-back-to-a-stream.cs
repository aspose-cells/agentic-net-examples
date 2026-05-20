using System;
using System.IO;
using Aspose.Cells;

class XmlMapExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the Excel file into a byte array
            byte[] excelBytes = File.ReadAllBytes(inputPath);

            // Load the workbook from a memory stream
            using (MemoryStream inputStream = new MemoryStream(excelBytes))
            {
                Workbook workbook = new Workbook(inputStream);

                // XML schema (XSD) definition as a string
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // Write schema to a temporary XSD file (required by Aspose.Cells API)
                string tempXsdPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempXsdPath, xmlSchema);

                // Add the XML map using the temporary XSD file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "SampleMap";

                // Clean up temporary XSD file
                if (File.Exists(tempXsdPath))
                    File.Delete(tempXsdPath);

                // Save the modified workbook to a new memory stream
                using (MemoryStream outputStream = new MemoryStream())
                {
                    workbook.Save(outputStream, SaveFormat.Xlsx);
                    outputStream.Position = 0; // Reset for downstream reading

                    // Write the stream to a physical file
                    File.WriteAllBytes(outputPath, outputStream.ToArray());
                }
            }

            Console.WriteLine("Workbook processed and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}