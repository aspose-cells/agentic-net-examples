using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlImportDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Prepare a simple XML schema and save it to a temporary file
            string schemaContent = @"<?xml version='1.0' encoding='utf-8'?>
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Products'>
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name='Product' maxOccurs='unbounded'>
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name='Name' type='xs:string'/>
                              <xs:element name='Price' type='xs:decimal'/>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";
            string schemaPath = "ProductsSchema.xsd";
            File.WriteAllText(schemaPath, schemaContent);

            // Add the XML map to the workbook using the schema file
            int mapIndex = wb.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap xmlMap = wb.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "ProductsMap";

            // Link a cell to a specific XML element path
            // Here we link cell A1 to the first Product's Name element
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;
            cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Products/Product/Name");

            // Prepare sample XML data and save it to a temporary file
            string xmlData = @"<?xml version='1.0' encoding='utf-8'?>
                <Products>
                  <Product>
                    <Name>Laptop</Name>
                    <Price>999.99</Price>
                  </Product>
                  <Product>
                    <Name>Phone</Name>
                    <Price>699.99</Price>
                  </Product>
                </Products>";
            string xmlPath = "ProductsData.xml";
            File.WriteAllText(xmlPath, xmlData);

            // Import the XML data into the workbook starting at cell A1 of Sheet1
            wb.ImportXml(xmlPath, "Sheet1", 0, 0);

            // Save the workbook to an Excel file
            wb.Save("ProductsWorkbook.xlsx");

            // Clean up temporary files (optional)
            File.Delete(schemaPath);
            File.Delete(xmlPath);
        }
    }
}