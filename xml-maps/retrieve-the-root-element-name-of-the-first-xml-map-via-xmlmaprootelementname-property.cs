using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Define an XML schema (XSD) that will be used to create the XML map
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Data'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' maxOccurs='unbounded'>
                                                    <xs:complexType>
                                                        <xs:sequence>
                                                            <xs:element name='Name' type='xs:string'/>
                                                            <xs:element name='Value' type='xs:integer'/>
                                                        </xs:sequence>
                                                    </xs:complexType>
                                                </xs:element>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file because Aspose.Cells expects a file path
            string tempXsdPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
            File.WriteAllText(tempXsdPath, xmlSchema);

            // Ensure the temporary XSD file exists before adding the XML map
            if (!File.Exists(tempXsdPath))
                throw new FileNotFoundException("Temporary XSD file was not created.", tempXsdPath);

            // Add the XML map to the workbook using the temporary XSD file
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);

            // Retrieve the first XmlMap from the collection
            XmlMap firstMap = workbook.Worksheets.XmlMaps[0];

            // Get the root element name using the RootElementName property
            string rootElementName = firstMap.RootElementName;

            // Output the root element name
            Console.WriteLine("Root Element Name: " + rootElementName);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            // Clean up the temporary XSD file if it exists
            try
            {
                string tempXsdPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
                if (File.Exists(tempXsdPath))
                    File.Delete(tempXsdPath);
            }
            catch
            {
                // Suppress any cleanup exceptions
            }
        }
    }
}