using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapPreserveWhitespaceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Fill sample data
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["B2"].PutValue("John Doe");

                // Define a simple XML schema for the map
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Employee'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='ID' type='xs:int'/>
                                                    <xs:element name='Name' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // NOTE: PreserveWhitespace property is not available in this version of Aspose.Cells.
                // If needed, use the appropriate API for the version you are targeting.

                // Prepare save options for XML export
                XmlSaveOptions saveOptions = new XmlSaveOptions
                {
                    XmlMapName = xmlMap.Name,
                    SheetNameAsElementName = true
                };

                // Save the workbook as an XML file
                string outputPath = "EmployeeData.xml";
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    class Program
    {
        static void Main(string[] args)
        {
            XmlMapPreserveWhitespaceDemo.Run();
        }
    }
}