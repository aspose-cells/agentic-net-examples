using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    public class Program
    {
        public static void Main()
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

            // Write the schema to a temporary file
            string tempXsdPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
            File.WriteAllText(tempXsdPath, xmlSchema);

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Ensure that at least one XmlMap exists
            if (workbook.Worksheets.XmlMaps.Count > 0 && xmlMap != null)
            {
                // Retrieve the root element name of the XML map
                string rootElementName = xmlMap.RootElementName;

                // Display the result
                Console.WriteLine("Root Element Name: " + rootElementName);
            }
            else
            {
                Console.WriteLine("No XML maps are defined in the workbook.");
            }

            // Optional: save the workbook
            // workbook.Save("XmlMapDemo.xlsx", SaveFormat.Xlsx);

            // Clean up temporary file
            if (File.Exists(tempXsdPath))
            {
                File.Delete(tempXsdPath);
            }
        }
    }
}