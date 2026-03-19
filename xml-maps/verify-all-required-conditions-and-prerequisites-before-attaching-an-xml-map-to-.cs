using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsXmlMapPrereqDemo
{
    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Verify prerequisite: license must be set for full functionality
            if (!workbook.IsLicensed)
            {
                Console.WriteLine("Warning: Workbook is not licensed. XML mapping may be limited.");
            }

            // 3. Verify prerequisite: the workbook should not contain macros when adding an XML map
            if (workbook.HasMacro)
            {
                Console.WriteLine("Error: Workbook contains macros. Remove macros before attaching an XML map.");
                return;
            }

            // 4. Prepare XML schema (must be a non‑empty string)
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            if (string.IsNullOrWhiteSpace(xmlSchema))
            {
                Console.WriteLine("Error: XML schema is empty or null.");
                return;
            }

            // 5. Verify that a map with the intended name does not already exist
            string intendedMapName = "MyXmlMap";
            foreach (XmlMap existingMap in workbook.Worksheets.XmlMaps)
            {
                if (existingMap.Name.Equals(intendedMapName, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Error: An XML map with the name '{intendedMapName}' already exists.");
                    return;
                }
            }

            // 6. Write schema to a temporary XSD file and attach the XML map
            string tempXsdPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xsd");
            File.WriteAllText(tempXsdPath, xmlSchema, Encoding.UTF8);

            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = intendedMapName; // set a friendly name

            // Clean up temporary file
            File.Delete(tempXsdPath);

            Console.WriteLine($"XML map '{xmlMap.Name}' added successfully at index {mapIndex}.");

            // 7. Save the workbook
            string outputPath = "WorkbookWithXmlMap.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}