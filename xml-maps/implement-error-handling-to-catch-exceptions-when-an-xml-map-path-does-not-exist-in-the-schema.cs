using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapQueryErrorHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (already present)
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample XML schema (XSD) defining a simple structure
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='ExistingElement' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Add the XML map to the workbook using the schema
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "DemoMap";

                // Link a cell to an existing element (optional, just for completeness)
                worksheet.Cells["A1"].PutValue("Sample");
                worksheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/ExistingElement");

                // Define a path that does NOT exist in the schema
                string invalidPath = "/Root/NonExistingElement";

                // Query the worksheet for cell areas mapped to the invalid path
                ArrayList cellAreas = worksheet.XmlMapQuery(invalidPath, xmlMap);

                // If the path is not present, the returned list will be empty
                if (cellAreas.Count == 0)
                {
                    Console.WriteLine($"No cells are mapped to the path '{invalidPath}'.");
                }
                else
                {
                    // This block would execute only if the path somehow returned results
                    CellArea area = (CellArea)cellAreas[0];
                    Console.WriteLine($"Found mapping at Row {area.StartRow}, Column {area.StartColumn}");
                }

                // Save the workbook (optional, demonstrates lifecycle rule usage)
                string outputPath = "XmlMapQueryErrorHandlingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (CellsException ex)
            {
                // Handle specific Aspose.Cells exceptions (e.g., invalid path)
                Console.WriteLine($"Aspose.Cells exception caught: {ex.Message}");
                Console.WriteLine($"Exception Type Code: {ex.Code}");
            }
            catch (Exception ex)
            {
                // Handle any other unexpected exceptions
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            XmlMapQueryErrorHandlingDemo.Run();
        }
    }
}