// Title: Add a Custom‑Named XML Map to a Workbook with Aspose.Cells (C#)
// Description: Demonstrates how to create a new Workbook, define an inline XML schema, add it to the Worksheets.XmlMaps collection using XmlMapCollection.Add, assign a custom name to the resulting XmlMap, and save the file as WorkbookWithXmlMap.xlsx.
// Keywords: Aspose.Cells XML map | XmlMapCollection Add | custom XmlMap name | C# Aspose.Cells example | add XML schema to workbook | set XmlMap.Name | export workbook with XML map | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells add XML map C# | how to set custom name for XmlMap | XmlMapCollection Add method usage | save workbook after adding XML map | C# example for XML schema mapping in Aspose.Cells
// Developer Intent: Create an XML map in a workbook and give it a user‑defined name using Aspose.Cells for .NET.
// Use Cases: Map XML data to worksheet cells with a recognizable map identifier. | Reuse a named XML map across multiple workbooks for consistent data import/export. | Prepare a workbook for automated XML export where the map name conveys the data structure.
// AI Prompts: Generate C# code that loads an XML schema from a file, adds it to a workbook with XmlMapCollection.Add, and sets a custom map name. | Show how to retrieve an existing XmlMap by name and modify its properties in Aspose.Cells. | Provide robust error handling for invalid XML schemas when adding an XmlMap in a .NET application.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new Workbook, define an inline XML schema, add it to the Worksheets.XmlMaps collection using XmlMapCollection.Add, assign a custom name to the resulting XmlMap, and save the file as WorkbookWithXmlMap.xlsx.
    public class AddXmlMapWithCustomName
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the XmlMapCollection from the workbook
                XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

                // Define an XML schema (inline string)
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Item' type='xs:string'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                     </xs:schema>";

                // Add the XML map to the collection
                int mapIndex = xmlMaps.Add(xmlSchema);

                // Retrieve the added XmlMap and assign a custom name
                XmlMap xmlMap = xmlMaps[mapIndex];
                xmlMap.Name = "MyCustomMap";

                // Save the workbook
                workbook.Save("WorkbookWithXmlMap.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddXmlMapWithCustomName.Run();
        }
    }
}
