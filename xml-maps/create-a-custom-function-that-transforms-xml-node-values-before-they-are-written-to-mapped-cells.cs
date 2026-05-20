using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsXmlTransformDemo
{
    // Custom class that overrides ExportTableOptions to transform cell values during export
    public class MyExportTableOptions : ExportTableOptions
    {
        // Example transformation: trim whitespace and convert strings to upper case
        private string TransformValue(string original)
        {
            if (original == null) return null;
            return original.Trim().ToUpperInvariant();
        }

        // Called for each cell when the workbook is exported.
        // Replaces the cell value if it is a string.
        public override bool PreprocessExportedValue(int cellRow, int cellColumn, CellValue value)
        {
            if (value.Type == CellValueType.IsString && value.Value is string str)
            {
                string transformed = TransformValue(str);
                value.Value = transformed;
                value.Type = CellValueType.IsString;
                return true; // value replaced
            }
            return false; // no change
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook workbook = new Workbook();

                // 2. Add an XML map (sample schema)
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
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

                // Add the schema from a string (isFile = false). Use overload that accepts only the schema string
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "ProductsMap";

                // 3. Link worksheet cells to the XML map paths
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;
                cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Products/Product/Name");   // A1
                cells.LinkToXmlMap(xmlMap.Name, 0, 1, "/Products/Product/Price"); // B1

                // 4. Prepare sample XML data
                string xmlData = @"<Products>
                                     <Product>
                                       <Name>  Laptop </Name>
                                       <Price>999.99</Price>
                                     </Product>
                                     <Product>
                                       <Name>phone</Name>
                                       <Price>699.50</Price>
                                     </Product>
                                   </Products>";

                // 5. Load the XML data into the workbook using XmlLoadOptions (mapping enabled)
                XmlLoadOptions loadOptions = new XmlLoadOptions
                {
                    IsXmlMap = true   // Enable mapping of XML to the linked cells
                };
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xmlData)))
                {
                    workbook.ImportXml(ms, sheet.Name, 0, 0);
                }

                // 6. Apply custom transformation by iterating cells and using ExportTableOptions logic
                MyExportTableOptions exportOptions = new MyExportTableOptions();

                for (int row = 0; row <= sheet.Cells.MaxDataRow; row++)
                {
                    for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell.Type == CellValueType.IsString)
                        {
                            CellValue cv = new CellValue
                            {
                                Type = CellValueType.IsString,
                                Value = cell.StringValue
                            };

                            if (exportOptions.PreprocessExportedValue(row, col, cv))
                            {
                                cell.PutValue(cv.Value);
                            }
                        }
                    }
                }

                // 7. Save the workbook (Excel format) to verify the transformed values
                string excelPath = "TransformedProducts.xlsx";
                workbook.Save(excelPath);

                // 8. Export the XML again to see the effect in the XML output
                string xmlPath = "TransformedProducts.xml";
                try
                {
                    workbook.ExportXml(xmlMap.Name, xmlPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"XML export failed: {ex.Message}");
                }

                Console.WriteLine($"Processing completed. Files saved: '{excelPath}', '{xmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}