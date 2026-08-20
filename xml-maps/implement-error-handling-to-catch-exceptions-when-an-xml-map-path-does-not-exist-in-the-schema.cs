// Title: C# Aspose.Cells: Catch CellsException for Invalid XmlMapQuery Paths
// Description: Demonstrates how to add an XML map from an XSD, link cells, and safely query a non‑existent XML path using nested try‑catch blocks. The example captures Aspose.Cells CellsException and generic errors, logs clear messages, and still saves the workbook.
// Keywords: Aspose.Cells XmlMapQuery error handling | C# catch CellsException | invalid XML map path | Aspose.Cells XML map exception | .NET XML map query | Aspose.Cells workbook save after error | handle missing XML element Aspose | XmlMapQuery try catch example
// Common Searches: Aspose.Cells XmlMapQuery invalid path exception | C# how to catch CellsException for XmlMapQuery | error handling for missing XML map element in Aspose.Cells | sample code for XmlMapQuery try catch | Aspose.Cells query non‑existent XML node
// Developer Intent: Add robust try‑catch logic to manage errors when an XmlMapQuery path is absent in the schema.
// Use Cases: Log a specific error when a queried XML element is not defined and continue processing. | Provide a fallback value or alternative mapping when XmlMapQuery returns no linked cells. | Ensure the workbook is saved and the application remains stable after handling a CellsException.
// AI Prompts: Create C# code that queries an Aspose.Cells XML map and includes separate catch blocks for CellsException and generic exceptions. | Show how to pre‑validate an XML map path before calling XmlMapQuery and return a custom warning if the path is missing. | Write a reusable helper method that wraps XmlMapQuery, catches CellsException, logs details, and returns an empty list when the path does not exist.

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to add an XML map from an XSD, link cells, and safely query a non‑existent XML path using nested try‑catch blocks. The example captures Aspose.Cells CellsException and generic errors, logs clear messages, and still saves the workbook.
    public class XmlMapQueryErrorHandlingDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample XML schema (XSD) defining a simple structure
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string' />
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Add the XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "SimpleMap";

                // Link a cell to a valid XML path (for demonstration)
                worksheet.Cells["A1"].PutValue("Sample");
                worksheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/Root/Item");

                // Attempt to query a non‑existent path in the XML map
                string invalidPath = "/Root/NonExistentElement";

                try
                {
                    // XmlMapQuery returns a list of CellArea objects; if the path is invalid,
                    // Aspose.Cells may throw a CellsException. We catch it to handle the error gracefully.
                    ArrayList cellAreas = worksheet.XmlMapQuery(invalidPath, xmlMap);

                    if (cellAreas.Count == 0)
                    {
                        Console.WriteLine($"No cells are linked to the path '{invalidPath}'.");
                    }
                    else
                    {
                        // This block would execute only if the path somehow exists.
                        CellArea area = (CellArea)cellAreas[0];
                        Console.WriteLine($"Found linked cell at Row {area.StartRow}, Column {area.StartColumn}.");
                    }
                }
                catch (CellsException ex)
                {
                    // Handle specific Aspose.Cells exceptions
                    Console.WriteLine($"Error querying XML map path '{invalidPath}': {ex.Message}");
                    Console.WriteLine($"Exception Type Code: {ex.Code}");
                }
                catch (Exception ex)
                {
                    // Handle any other unexpected exceptions
                    Console.WriteLine($"Unexpected error: {ex.Message}");
                }

                // Save the workbook (optional, demonstrates normal lifecycle usage)
                workbook.Save("XmlMapQueryErrorHandlingDemo.xlsx");
                Console.WriteLine("Workbook saved as XmlMapQueryErrorHandlingDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
