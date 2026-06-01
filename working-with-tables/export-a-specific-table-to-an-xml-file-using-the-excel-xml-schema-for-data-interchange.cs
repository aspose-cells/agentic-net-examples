using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

namespace AsposeCellsExportTableToXml
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that will be part of the table
                sheet.Cells["A1"].PutValue("Id");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue("Charlie");

                // Add a table (ListObject) covering the data range A1:B4
                int firstRow = 0;      // zero‑based index for row 1
                int firstColumn = 0;   // zero‑based index for column A
                int totalRows = 4;
                int totalColumns = 2;

                // Add returns the index of the created ListObject
                int tableIdx = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
                ListObject table = sheet.ListObjects[tableIdx];
                table.DisplayName = "PeopleTable";

                // Define a simple XML schema that matches the table structure
                string xmlSchema = @"
                    <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                      <xs:element name='People'>
                        <xs:complexType>
                          <xs:sequence>
                            <xs:element name='Person' maxOccurs='unbounded'>
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

                // Add the XML map to the workbook using the schema
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "PeopleMap";

                // Configure XML save options to export only the table range
                XmlSaveOptions saveOptions = new XmlSaveOptions
                {
                    ExportArea = new CellArea
                    {
                        StartRow = table.DataRange.FirstRow,
                        EndRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1,
                        StartColumn = table.DataRange.FirstColumn,
                        EndColumn = table.DataRange.FirstColumn + table.DataRange.ColumnCount - 1
                    },
                    XmlMapName = xmlMap.Name,
                    SheetNameAsElementName = true,
                    DataAsAttribute = false
                };

                // Determine output path and ensure directory exists
                string outputPath = "PeopleTableExport.xml";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an XML file using the configured options
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Table exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}