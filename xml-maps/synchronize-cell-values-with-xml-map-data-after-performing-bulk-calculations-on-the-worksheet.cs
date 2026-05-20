using System;
using System.IO;
using System.Text;
using Aspose.Cells;

class SyncXmlMapDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string xmlOutputPath = "output.xml";

            // Load existing workbook or create a new one if it doesn't exist.
            Workbook wb;
            if (File.Exists(inputPath))
            {
                wb = new Workbook(inputPath);
            }
            else
            {
                wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
            }

            // Access the first worksheet and its cells.
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // ----- Bulk calculations -----
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Calculate formulas.
            wb.CalculateFormula();

            // ----- XML map setup -----
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Value' type='xs:string'/>
                                                <xs:element name='Total' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                </xs:schema>";

            // Write schema to a temporary file because Aspose.Cells expects a file path.
            string tempSchemaPath = Path.GetTempFileName();
            File.WriteAllText(tempSchemaPath, xmlSchema, Encoding.UTF8);

            // Add the XML map from the temporary schema file.
            int mapIndex = wb.Worksheets.XmlMaps.Add(tempSchemaPath);
            // Clean up the temporary file.
            File.Delete(tempSchemaPath);

            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "DemoMap";

            // Link worksheet cells to the XML map paths.
            cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Value"); // A1 -> /Root/Value
            cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Root/Total"); // B1 -> /Root/Total

            // Re‑calculate to ensure linked cells contain the latest values.
            wb.CalculateFormula();

            // ----- Export synchronized XML -----
            wb.ExportXml(xmlMap.Name, xmlOutputPath);

            // Save the workbook.
            wb.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}