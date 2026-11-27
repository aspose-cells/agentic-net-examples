// Title: Create and use an XML map from an in‑memory schema string in Aspose.Cells for .NET
// Common Searches: Aspose.Cells add xml map from schema string c# | load xml map without a file using Workbook.Worksheets.XmlMaps.Add | export worksheet to xml using an in‑memory xml map Aspose.Cells | dynamic xml schema generation for Excel export with Aspose.Cells .NET | map worksheet data to xml using a string schema Aspose.Cells
// Tags: Workbook.Worksheets.XmlMaps.Add schema string | in‑memory XML map Aspose.Cells .NET | export worksheet to XML with XmlMap | dynamic XML schema generation for Excel | add xml map using schema string Aspose.Cells

using System;
using Aspose.Cells;

namespace XmlMapFromStringDemo
{
    // The example creates a new Workbook, fills it with sample employee data, defines an XML schema as a string, adds the schema to the workbook via Workbook.Worksheets.XmlMaps.Add, assigns a friendly name to the XmlMap, exports the worksheet data to an XML file using ExportXml, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // 2. Populate the worksheet with sample data that matches the XML schema.
                // Header row
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["C1"].PutValue("Salary");

                // Data rows
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("John Doe");
                worksheet.Cells["C2"].PutValue(50000);

                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Jane Smith");
                worksheet.Cells["C3"].PutValue(62000);

                // 3. Define the XML schema as a string.
                // This schema describes an <Employees> root element containing multiple <Employee> elements.
                string xmlSchema = @"
<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Employees'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Employee' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='ID' type='xs:int'/>
              <xs:element name='Name' type='xs:string'/>
              <xs:element name='Salary' type='xs:decimal'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                // 4. Add the XML map to the workbook directly from the string.
                // The Add method of XmlMapCollection accepts a string containing the schema.
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // 5. Optionally give the map a friendly name (useful when exporting).
                xmlMap.Name = "EmployeesMap";

                // 6. Export the worksheet data to an XML file using the map.
                // The ExportXml method requires the map name and the output file path.
                string outputXmlPath = "EmployeesExport.xml";
                workbook.ExportXml(xmlMap.Name, outputXmlPath);

                Console.WriteLine($"XML map '{xmlMap.Name}' added successfully.");
                Console.WriteLine($"Root element name detected: {xmlMap.RootElementName}");
                Console.WriteLine($"Data exported to: {outputXmlPath}");

                // 7. Save the workbook itself for verification (optional).
                string workbookPath = "WorkbookWithXmlMap.xlsx";
                workbook.Save(workbookPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                // Any exception will be printed to the console – no hidden failures.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
