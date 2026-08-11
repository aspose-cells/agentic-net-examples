// Title: Rename an XML Map and Save the Workbook with XmlSaveOptions in Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, write a simple XSD to a temporary file, add it as an XML map, rename the map, configure XmlSaveOptions with the new map name, and save the workbook so the modified XML map is persisted in the output XML file.
// Keywords: Aspose.Cells XML map rename | XmlSaveOptions C# | save workbook with XML map | temporary XSD Aspose.Cells | persist XML map changes | export workbook to XML | C# Aspose.Cells example
// Common Searches: how to rename an XML map in Aspose.Cells | save workbook with modified XML map C# | XmlSaveOptions usage Aspose.Cells | add XML map from XSD file Aspose.Cells | persist XML map after editing
// Developer Intent: Persist modifications to an XML map by saving the workbook with the appropriate XmlSaveOptions.
// Use Cases: Rename an existing XML map and export only that map to an XML file. | Create a workbook, attach an XML map from a schema file, change its properties, and save for downstream XML processing. | Generate a temporary XSD, bind it as an XML map, adjust settings, and ensure the configuration is stored in the saved workbook.
// AI Prompts: Show how to add an XML map from an in‑memory XSD string without writing a temporary file using Aspose.Cells. | Provide code to update multiple XML maps in a workbook and save each with distinct XmlSaveOptions. | Explain how to verify that a renamed XML map is correctly included in the exported XML file.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapSaveDemo
{
    // Demonstrates how to create a workbook, write a simple XSD to a temporary file, add it as an XML map, rename the map, configure XmlSaveOptions with the new map name, and save the workbook so the modified XML map is persisted in the output XML file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Define a simple XML schema for the map
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

                // Ensure the temporary schema file exists before adding the XML map
                if (!File.Exists(tempSchemaPath))
                    throw new FileNotFoundException("Temporary XML schema file was not created.", tempSchemaPath);

                // Add the XML map to the workbook using the temporary schema file
                int mapIndex = workbook.Worksheets.XmlMaps.Add(tempSchemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Modify the XML map (e.g., change its name)
                xmlMap.Name = "MyCustomXmlMap";

                // Prepare save options to include the modified XML map in the saved file
                XmlSaveOptions saveOptions = new XmlSaveOptions
                {
                    XmlMapName = xmlMap.Name // ensure the map is exported
                };

                // Save the workbook, persisting the XML map changes
                string outputPath = "WorkbookWithModifiedXmlMap.xml";
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved successfully with modified XML map at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
