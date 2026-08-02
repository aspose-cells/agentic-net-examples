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

            // Define a simple XML schema (XSD) as a string
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write the schema to a temporary file (Aspose.Cells expects a file path)
            string tempXsdPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
            File.WriteAllText(tempXsdPath, xmlSchema);

            // Ensure the temporary XSD file exists before adding the XML map
            if (!File.Exists(tempXsdPath))
                throw new FileNotFoundException("Temporary XSD file not found.", tempXsdPath);

            // Add the XML map to the workbook; the method returns the index of the added map
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);

            // Retrieve the root element name of the map using the index
            string rootElementName = workbook.Worksheets.XmlMaps[mapIndex].RootElementName;

            // Display the result
            Console.WriteLine("Root Element Name: " + rootElementName);

            // Save the workbook (optional, demonstrates lifecycle usage)
            string outputPath = "DemoOutput.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}