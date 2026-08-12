// Title: Update Excel cells linked to an XML map and export the modified XML – Aspose.Cells C# example
// Description: Demonstrates how to create a workbook, add an XML map from a temporary XSD, bind cells A1 and B1 to /Root/Title and /Root/Date, set initial values, modify the cells, and export the XML before and after the changes. The workbook can also be saved for further use.
// Keywords: Aspose.Cells XML map C# | link Excel cells to XML elements | export updated XML from workbook | modify mapped cell values | XML map synchronization | C# Excel to XML conversion | temporary XSD file Aspose
// Common Searches: how to update cells linked to an XML map using Aspose.Cells | C# export XML after editing mapped Excel cells | Aspose.Cells example for XML map data binding | change Excel cell values and reflect them in XML with .NET | add XML map to workbook programmatically
// Developer Intent: Programmatically change the values of cells that are bound to an XML map and have those changes automatically written back to the source XML document.
// Use Cases: Populate an Excel template from XML, let users edit the linked cells, then generate an updated XML file for downstream processing. | Maintain configuration data in XML while providing a spreadsheet UI for non‑technical users to edit values safely. | Automate report generation by programmatically adjusting mapped cells and exporting the resulting XML for integration with other services.
// AI Prompts: Write C# code that loads an XSD, adds it as an XML map to a workbook, links specific cells to XML nodes, updates those cells, and exports the revised XML using Aspose.Cells. | Explain how to retrieve the index of a newly added XML map and how to format DateTime values for xs:date elements when linking cells. | Provide robust error‑handling patterns for ExportXml when the output path is invalid or the specified XML map does not exist.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapUpdateDemo
{
    // Demonstrates how to create a workbook, add an XML map from a temporary XSD, bind cells A1 and B1 to /Root/Title and /Root/Date, set initial values, modify the cells, and export the XML before and after the changes. The workbook can also be saved for further use.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Define a simple XML schema that will be used as the map
                string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                        <xs:element name='Root'>
                                            <xs:complexType>
                                                <xs:sequence>
                                                    <xs:element name='Title' type='xs:string'/>
                                                    <xs:element name='Date' type='xs:date'/>
                                                </xs:sequence>
                                            </xs:complexType>
                                        </xs:element>
                                    </xs:schema>";

                // Write the schema to a temporary XSD file (required by Aspose.Cells API)
                string tempXsdPath = Path.Combine(Path.GetTempPath(), "DemoMap.xsd");
                File.WriteAllText(tempXsdPath, xmlSchema);

                // Add the XML map to the workbook using the XSD file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempXsdPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "DemoMap";

                // Get the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Link cells to XML elements using the map
                // A1 -> /Root/Title
                // B1 -> /Root/Date
                cells.LinkToXmlMap("DemoMap", 0, 0, "/Root/Title");
                cells.LinkToXmlMap("DemoMap", 0, 1, "/Root/Date");

                // Set initial values in the linked cells
                cells["A1"].PutValue("Initial Title");
                cells["B1"].PutValue(new DateTime(2023, 1, 1));

                // Export the XML to see the initial state
                workbook.ExportXml("InitialOutput.xml", "DemoMap");

                // Update the cell values – these changes will be reflected in the XML map
                cells["A1"].PutValue("Updated Title");
                cells["B1"].PutValue(new DateTime(2024, 12, 31));

                // Export the XML again; the file now contains the updated values
                workbook.ExportXml("UpdatedOutput.xml", "DemoMap");

                // Save the workbook (optional, just to keep the Excel file)
                workbook.Save("MappedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
