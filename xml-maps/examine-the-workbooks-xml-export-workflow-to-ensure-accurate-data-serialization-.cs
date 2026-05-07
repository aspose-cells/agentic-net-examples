using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlExportDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook if it exists; otherwise create a new one with sample data.
            string workbookPath = "Book1.xlsx";
            Workbook wb;
            if (File.Exists(workbookPath))
            {
                wb = new Workbook(workbookPath);
            }
            else
            {
                wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue("Id");
                ws.Cells["B1"].PutValue("Name");
                ws.Cells["A2"].PutValue(1);
                ws.Cells["B2"].PutValue("Alice");
                ws.Cells["A3"].PutValue(2);
                ws.Cells["B3"].PutValue("Bob");
            }

            // Ensure at least one XML map exists; if not, create a simple schema and add the map.
            if (wb.Worksheets.XmlMaps.Count == 0)
            {
                string simpleSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
  <xs:element name='Root'>
    <xs:complexType>
      <xs:sequence>
        <xs:element name='Row' maxOccurs='unbounded'>
          <xs:complexType>
            <xs:sequence>
              <xs:element name='Id' type='xs:int'/>
              <xs:element name='Name' type='xs:string'/>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";

                // Write schema to a temporary file because Add(string) expects a file path.
                string tempSchemaPath = Path.GetTempFileName();
                File.WriteAllText(tempSchemaPath, simpleSchema);

                try
                {
                    int mapIndex = wb.Worksheets.XmlMaps.Add(tempSchemaPath);
                    XmlMap map = wb.Worksheets.XmlMaps[mapIndex];
                    map.Name = "SimpleMap";
                }
                finally
                {
                    // Clean up the temporary file.
                    if (File.Exists(tempSchemaPath))
                        File.Delete(tempSchemaPath);
                }
            }

            // Retrieve the first XML map to use for export.
            XmlMap exportMap = wb.Worksheets.XmlMaps[0];
            string xmlMapName = exportMap.Name;

            // Export the XML data linked to the map into a memory stream.
            using (MemoryStream xmlStream = new MemoryStream())
            {
                wb.ExportXml(xmlMapName, xmlStream);
                xmlStream.Position = 0; // Reset for reading.

                // Display the exported XML content.
                using (StreamReader reader = new StreamReader(xmlStream))
                {
                    string xmlContent = reader.ReadToEnd();
                    Console.WriteLine("Exported XML Content:");
                    Console.WriteLine(xmlContent);
                }

                // Save the stream to a physical file.
                string outputPath = "ExportedData.xml";
                File.WriteAllBytes(outputPath, xmlStream.ToArray());
                Console.WriteLine($"XML saved to {outputPath}");
            }

            // Save the entire workbook as an XML file using XmlSaveOptions, specifying the map.
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                XmlMapName = xmlMapName,
                SheetNameAsElementName = true,
                DataAsAttribute = false
            };
            string xmlWorkbookPath = "WorkbookWithMap.xml";
            wb.Save(xmlWorkbookPath, saveOptions);
            Console.WriteLine($"Workbook saved as XML with map using {xmlWorkbookPath}");
        }
    }
}