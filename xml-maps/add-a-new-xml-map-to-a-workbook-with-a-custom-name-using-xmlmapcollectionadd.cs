// Title: Add a Named XML Map to a Workbook with Aspose.Cells (C#) using XmlMapCollection.Add
// Description: Demonstrates how to create a new Workbook, write an XSD schema to a temporary file, add the schema as an XML map with XmlMaps.Add, assign a custom Name to the XmlMap, and save the workbook containing the named map.
// Keywords: Aspose.Cells C# XML map | XmlMapCollection.Add | custom XmlMap name | add XSD schema to workbook | temporary XSD file Aspose.Cells | save workbook with XML map | .NET Excel XML mapping
// Common Searches: Aspose.Cells add XML map with custom name C# | XmlMapCollection.Add example .NET | set XmlMap.Name after adding to workbook | how to use temporary XSD file with Aspose.Cells | save Excel file with XML map Aspose
// Developer Intent: Create an XML map in a workbook and give it a custom identifier.
// Use Cases: Import XML data by first adding its XSD as a named map, then referencing the map for data import. | Build a reusable Excel template that includes a predefined, meaningfully‑named XML map for downstream XML export. | Generate workbooks with multiple distinct XML maps, each labeled with a unique name for easy programmatic access.
// AI Prompts: Write C# code that adds several XML maps from different XSD files to a workbook, assigning each a unique custom name using Aspose.Cells. | Explain how to replace the XSD schema of an existing XmlMap while keeping its custom Name property intact. | Provide a step‑by‑step guide to validate an XML file against a previously added XmlMap before importing its data into a worksheet.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a new Workbook, write an XSD schema to a temporary file, add the schema as an XML map with XmlMaps.Add, assign a custom Name to the XmlMap, and save the workbook containing the named map.
class AddXmlMapDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define an XML schema as a string
            string xmlSchema = @"<xs:schema xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                    <xs:element name='Root'>
                                        <xs:complexType>
                                            <xs:sequence>
                                                <xs:element name='Item' type='xs:string'/>
                                            </xs:sequence>
                                        </xs:complexType>
                                    </xs:element>
                                 </xs:schema>";

            // Write the schema to a temporary file (Aspose.Cells expects a file path)
            string tempSchemaPath = Path.Combine(Path.GetTempPath(), "tempSchema.xsd");
            File.WriteAllText(tempSchemaPath, xmlSchema);

            // Ensure the temporary schema file exists before adding
            if (!File.Exists(tempSchemaPath))
                throw new FileNotFoundException("Temporary XML schema file was not created.", tempSchemaPath);

            // Add the XML map to the workbook; Add returns the map index
            int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);

            // Retrieve the newly added XmlMap and set a custom name
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
            xmlMap.Name = "MyCustomMap";

            // Save the workbook with the XML map attached
            string outputPath = "WorkbookWithXmlMap.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
