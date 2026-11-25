// Title: How to delete an XML map from an Excel workbook by index using Aspose.Cells XmlMaps.RemoveAt in C#
// Common Searches: C# Aspose.Cells remove XML map by index from .xlsx file | Delete unwanted XML schema mapping in Excel using Aspose.Cells XmlMaps.RemoveAt | Remove specific XML map from workbook programmatically with Aspose.Cells | Aspose.Cells XmlMaps.RemoveAt example for removing XML map in C# | How to remove an XML map from an Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells XmlMaps.RemoveAt usage | C# delete XML map in Excel workbook | Excel XML schema mapping removal with Aspose.Cells | programmatic XML map management Aspose.Cells | example removing XML map from .xlsx file

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveXmlMapDemo
{
    // The sample creates a workbook, adds sample data and an XML schema as an XmlMap, saves the file, then removes the XmlMap by its index using Workbook.Worksheets.XmlMaps.RemoveAt and saves the workbook again, demonstrating how to delete an unwanted XML map in C# with Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to save the workbook before and after removal
            string initialPath = "WorkbookWithXmlMap.xlsx";
            string finalPath = "WorkbookWithoutXmlMap.xlsx";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet (optional, just to have content)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Define a simple XML schema (XSD) as a string.
            // This schema will be used to create an XmlMap.
            string xmlSchema = @"
                <xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                  <xs:element name='Employees'>
                    <xs:complexType>
                      <xs:sequence>
                        <xs:element name='Employee' maxOccurs='unbounded'>
                          <xs:complexType>
                            <xs:sequence>
                              <xs:element name='ID' type='xs:int'/>
                              <xs:element name='Name' type='xs:string'/>
                            </xs:sequence>
                          </xs:complexType>
                        </xs:element>
                      </xs:sequence>
                    </xs:complexType>
                  </xs:element>
                </xs:schema>";

            try
            {
                // Add the XML map to the workbook. The Add method returns the index of the new map.
                int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchema);
                Console.WriteLine($"XML map added at index {mapIndex}.");

                // Verify that the map exists
                if (workbook.Worksheets.XmlMaps.Count > 0)
                {
                    Console.WriteLine($"Total XML maps before removal: {workbook.Worksheets.XmlMaps.Count}");
                }
                else
                {
                    Console.WriteLine("No XML maps were added – unexpected state.");
                }

                // Save the workbook with the map so we can see the before state (optional)
                workbook.Save(initialPath);
                Console.WriteLine($"Workbook saved with XML map to '{initialPath}'.");

                // Remove the XML map by its index using RemoveAt
                workbook.Worksheets.XmlMaps.RemoveAt(mapIndex);
                Console.WriteLine($"XML map at index {mapIndex} removed.");

                // Confirm removal
                Console.WriteLine($"Total XML maps after removal: {workbook.Worksheets.XmlMaps.Count}");

                // Save the workbook after removal
                workbook.Save(finalPath);
                Console.WriteLine($"Workbook saved without XML map to '{finalPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
