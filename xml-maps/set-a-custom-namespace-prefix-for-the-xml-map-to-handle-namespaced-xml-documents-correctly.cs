using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapCustomNamespacePrefixDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["B2"].PutValue("John Doe");

                // Define an XML schema that uses a custom namespace prefix (e.g., "emp")
                string xmlSchema = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
                               xmlns:emp='http://schemas.example.com/employee'
                               targetNamespace='http://schemas.example.com/employee'
                               elementFormDefault='qualified'>
                        <xs:element name='Employee' type='emp:EmployeeType'/>
                        <xs:complexType name='EmployeeType'>
                            <xs:sequence>
                                <xs:element name='ID' type='xs:int'/>
                                <xs:element name='Name' type='xs:string'/>
                            </xs:sequence>
                        </xs:complexType>
                    </xs:schema>";

                // Add the XML map to the workbook; the map will retain the namespace prefix defined above
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Assign a friendly name to the map (optional)
                xmlMap.Name = "EmployeeMap";

                // Export the worksheet data to an XML file using the map.
                // The exported XML will contain the custom namespace prefix "emp".
                string exportPath = "EmployeeOutput.xml";
                workbook.ExportXml(xmlMap.Name, exportPath);

                // Save the workbook (optional, to keep the map definition inside the file)
                workbook.Save("EmployeeWorkbook.xlsx");

                Console.WriteLine("XML exported with custom namespace prefix to: " + exportPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            XmlMapCustomNamespacePrefixDemo.Run();
        }
    }
}