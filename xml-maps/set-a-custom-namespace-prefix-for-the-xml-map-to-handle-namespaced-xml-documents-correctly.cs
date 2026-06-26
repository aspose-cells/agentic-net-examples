using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapCustomNamespacePrefixDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["B2"].PutValue("John Doe");

            // XML schema with a custom namespace prefix "ns"
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'
                           xmlns:ns='http://example.com/employee'
                           targetNamespace='http://example.com/employee'
                           elementFormDefault='qualified'>
                    <xs:element name='Employee' type='ns:EmployeeType'/>
                    <xs:complexType name='EmployeeType'>
                        <xs:sequence>
                            <xs:element name='ID' type='xs:int'/>
                            <xs:element name='Name' type='xs:string'/>
                        </xs:sequence>
                    </xs:complexType>
                </xs:schema>";

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Set a custom name for the map (used during export)
            xmlMap.Name = "ns:EmployeeMap";

            // Export the worksheet data to an XML file using the map name
            string exportPath = "EmployeeData.xml";

            // Ensure the directory for the export file exists
            string exportDir = Path.GetDirectoryName(Path.GetFullPath(exportPath));
            if (!Directory.Exists(exportDir))
            {
                Directory.CreateDirectory(exportDir);
            }

            workbook.ExportXml(xmlMap.Name, exportPath);

            // Save the workbook with XML mapping information
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                XmlMapName = xmlMap.Name // store the map name in the saved file
            };
            string workbookPath = "WorkbookWithXmlMap.xlsx";
            workbook.Save(workbookPath, saveOptions);

            Console.WriteLine("XML exported to: " + exportPath);
            Console.WriteLine("Workbook saved with XML map: " + workbookPath);
        }
    }
}