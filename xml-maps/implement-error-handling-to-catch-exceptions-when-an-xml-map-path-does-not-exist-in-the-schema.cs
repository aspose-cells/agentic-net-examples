using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapQueryErrorHandlingDemo
    {
        // Entry point required for console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define a simple XML schema (XSD) as a string
            string xsd = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                            <xs:element name='Root'>
                                <xs:complexType>
                                    <xs:sequence>
                                        <xs:element name='Item' type='xs:string' />
                                    </xs:sequence>
                                </xs:complexType>
                            </xs:element>
                          </xs:schema>";

            // Add the XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(xsd);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "SimpleMap";

            // Optionally link a cell to a valid XML path
            worksheet.Cells["A1"].PutValue("Sample");
            worksheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

            // Query a path that exists – should return a non‑empty list
            try
            {
                ArrayList areas = worksheet.XmlMapQuery("/Root/Item", xmlMap);
                Console.WriteLine($"Existing path query returned {areas.Count} area(s).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error querying existing path: {ex.Message}");
            }

            // Query a path that does NOT exist – demonstrate error handling
            try
            {
                // This path is not defined in the schema; Aspose.Cells may throw an exception
                ArrayList missingAreas = worksheet.XmlMapQuery("/Root/NonExisting", xmlMap);
                if (missingAreas.Count == 0)
                {
                    Console.WriteLine("Path not found in XML map – returned empty list.");
                }
                else
                {
                    Console.WriteLine($"Unexpectedly found {missingAreas.Count} area(s) for non‑existing path.");
                }
            }
            catch (CellsException cex)
            {
                // Handle Aspose.Cells specific exception
                Console.WriteLine($"CellsException caught: {cex.Message}");
                Console.WriteLine($"Exception Type Code: {cex.Code}");
            }
            catch (Exception ex)
            {
                // Handle any other exception
                Console.WriteLine($"General exception caught: {ex.Message}");
            }

            // Save the workbook (optional)
            try
            {
                string outputPath = "XmlMapQueryDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}