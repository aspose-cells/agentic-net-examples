// Title: How to update a cell in an XML‑mapped Excel table and export the modified XML using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update value in XML mapped ListObject cell | Export modified XML after changing Excel cell with Aspose.Cells .NET | Change product price in XML‑mapped worksheet using Aspose.Cells C# example | Programmatically edit XML map data in Excel and save updated XML with Aspose.Cells | C# code to import XML schema, bind to table, modify cell, and export XML using Aspose.Cells
// Tags: XML map cell update Aspose.Cells C# | export updated XML workbook Aspose.Cells | ListObject bound to XmlMap Aspose.Cells | import XML schema into worksheet Aspose.Cells | modify product price in Excel table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace XmlMapUpdateDemo
{
    // The example creates a Workbook, adds an XML schema as an XmlMap, imports XML data into Sheet1, creates a ListObject covering the imported range, finds the row with ID 2, updates its Price cell to 0.35, then exports the XmlMap to an updated XML file and saves the workbook as an XLSX.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for output files (adjust as needed)
            string outputXmlPath = Path.Combine(Environment.CurrentDirectory, "UpdatedData.xml");
            string outputXlsxPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithXmlMap.xlsx");

            // Sample XML schema (XSD) defining the structure of the XML data
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Products'>
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name='Product' maxOccurs='unbounded'>
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name='ID' type='xs:int'/>
                              <xs:element name='Name' type='xs:string'/>
                              <xs:element name='Price' type='xs:double'/>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";

            // Sample XML data that conforms to the above schema
            string xmlData = @"
                <Products>
                  <Product>
                    <ID>1</ID>
                    <Name>Apple</Name>
                    <Price>0.5</Price>
                  </Product>
                  <Product>
                    <ID>2</ID>
                    <Name>Banana</Name>
                    <Price>0.3</Price>
                  </Product>
                  <Product>
                    <ID>3</ID>
                    <Name>Cherry</Name>
                    <Price>1.2</Price>
                  </Product>
                </Products>";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and add the XML map (schema) to it
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();

                // Add the XML map using the schema string
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // ------------------------------------------------------------
                // 2. Import the XML data into the first worksheet using the map
                // ------------------------------------------------------------
                // The Workbook.ImportXml overload imports XML data and links it to the map.
                // The default sheet name for a new workbook is "Sheet1".
                workbook.ImportXml(xmlData, "Sheet1", 0, 0);

                // ------------------------------------------------------------
                // 3. Create a ListObject (table) that is bound to the imported XML map
                // ------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range after import
                int lastRow = sheet.Cells.MaxDataRow;
                int lastCol = sheet.Cells.MaxDataColumn;

                // Add a ListObject covering the whole used range (including header)
                int tableIndex = sheet.ListObjects.Add(0, 0, lastRow, lastCol, true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Verify that the ListObject is linked to our XmlMap
                Console.WriteLine($"ListObject '{table.DisplayName}' is linked to XmlMap: {table.XmlMap.Name}");

                // ------------------------------------------------------------
                // 4. Update a value in a mapped cell via the ListObject API
                // ------------------------------------------------------------
                // Example: Change the price of the product with ID = 2 (Banana) to 0.35
                int idColumnIndex = table.StartColumn;               // ID column
                int priceColumnIndex = table.StartColumn + 2;        // Price column (ID, Name, Price)

                int targetRow = -1;
                for (int r = table.StartRow + 1; r <= table.EndRow; r++) // skip header row
                {
                    Cell idCell = sheet.Cells[r, idColumnIndex];
                    if (idCell.IntValue == 2)
                    {
                        targetRow = r;
                        break;
                    }
                }

                if (targetRow != -1)
                {
                    // Update the price cell directly; the change will be reflected in the XML map
                    sheet.Cells[targetRow, priceColumnIndex].PutValue(0.35);
                    Console.WriteLine("Updated price for product ID 2 to 0.35.");
                }
                else
                {
                    Console.WriteLine("Product with ID 2 not found.");
                }

                // ------------------------------------------------------------
                // 5. Export the updated XML document – changes are reflected automatically
                // ------------------------------------------------------------
                workbook.ExportXml(xmlMap.Name, outputXmlPath);
                Console.WriteLine($"Updated XML exported to: {outputXmlPath}");

                // ------------------------------------------------------------
                // 6. Save the workbook (optional, for visual verification)
                // ------------------------------------------------------------
                workbook.Save(outputXlsxPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to: {outputXlsxPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
