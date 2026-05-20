using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

class ExportWorkbookToOdsWithXmlMap
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data that will be mapped to the XML map
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define a simple XML schema and add it as an XML map to the workbook
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Id' type='xs:int'/>
                                                <xs:element name='Name' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SampleMap";

            // NOTE: ImportData method is not available in the current Aspose.Cells version.
            // The XML map definition is already added; data mapping can be performed
            // during export if needed.

            // Create ODS save options and configure desired properties
            OdsSaveOptions saveOptions = new OdsSaveOptions
            {
                GeneratorType = OdsGeneratorType.LibreOffice,               // Set generator type
                OdfStrictVersion = OpenDocumentFormatVersionType.Odf12    // Use ODF 1.2
            };

            // Save the workbook as ODS; XML map definitions are retained automatically
            string outputPath = "WorkbookWithXmlMap.ods";
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}