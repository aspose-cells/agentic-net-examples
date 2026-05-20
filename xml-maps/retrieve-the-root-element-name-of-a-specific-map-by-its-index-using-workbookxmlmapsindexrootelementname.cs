using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRootElementDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a simple XML schema (XSD) with a root element named "Data"
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Data'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string' />
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Write the schema to a temporary file (Aspose.Cells expects a file path)
                string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
                File.WriteAllText(tempSchemaPath, xmlSchema);

                // Ensure the temporary schema file exists before adding the XML map
                if (!File.Exists(tempSchemaPath))
                    throw new FileNotFoundException("Temporary XML schema file was not created.", tempSchemaPath);

                // Add the XML map to the workbook; the method returns the index of the added map
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);

                // Retrieve the root element name of the map using its index
                string rootElementName = workbook.Worksheets.XmlMaps[mapIndex].RootElementName;

                // Display the retrieved root element name
                Console.WriteLine($"Root element name of map at index {mapIndex}: {rootElementName}");

                // Save the workbook
                string outputPath = "XmlMapRootElementDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");

                // Clean up the temporary schema file
                File.Delete(tempSchemaPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}