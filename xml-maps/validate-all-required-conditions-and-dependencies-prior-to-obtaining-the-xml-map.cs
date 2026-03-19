using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ValidateRootElementNameDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define a simple XML schema for the map
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
            string tempSchemaPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
            File.WriteAllText(tempSchemaPath, xmlSchema);

            try
            {
                // Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Validate the map
                if (xmlMap == null)
                {
                    Console.WriteLine("XmlMap is null. Cannot obtain root element name.");
                    return;
                }

                if (string.IsNullOrEmpty(xmlMap.Name))
                {
                    xmlMap.Name = "DefaultMap";
                }

                // Retrieve and display the root element name
                string rootElementName = xmlMap.RootElementName;
                if (string.IsNullOrEmpty(rootElementName))
                {
                    Console.WriteLine("Root element name is empty or undefined.");
                }
                else
                {
                    Console.WriteLine("Root Element Name: " + rootElementName);
                }

                // Save the workbook
                workbook.Save("ValidateRootElementNameDemo.xlsx");
            }
            finally
            {
                // Clean up temporary schema file
                if (File.Exists(tempSchemaPath))
                {
                    File.Delete(tempSchemaPath);
                }
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidateRootElementNameDemo.Run();
        }
    }
}