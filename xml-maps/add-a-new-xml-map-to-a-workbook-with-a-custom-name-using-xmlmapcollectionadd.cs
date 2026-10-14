// Title: Create and add a custom‑named XML map to a workbook with Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an XSD file, accesses the Workbook.XmlMaps collection via reflection, and adds a new XML map called 'MyCustomMap' before saving the workbook. | Show how to conditionally add an XML map only when the XmlMaps property exists in the current Aspose.Cells version, handling missing support gracefully. | Provide a complete example that creates a new Workbook, registers a custom XML map from a schema string, and writes the result to an .xlsx file.
// Common Searches: how to add a custom XML map to a workbook using Aspose.Cells C# | Aspose.Cells XmlMapCollection Add method with schema file example | reflection access XmlMaps property Aspose.Cells version check
// Tags: add xml map Aspose.Cells C# | xmlmapcollection add method | load xml schema workbook Aspose.Cells | reflection xmlmaps property check | save workbook with xml mapping

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, reads an XML schema file if present, uses reflection to obtain the XmlMaps collection, invokes its Add method to register a custom‑named XML map, and saves the workbook as an .xlsx file while safely handling environments where XML mapping is unavailable.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the XML schema if the file exists
            string schemaPath = "schema.xml";
            if (File.Exists(schemaPath))
            {
                string schemaXml = File.ReadAllText(schemaPath);

                // Add a new XML map with a custom name to the workbook
                // Note: XmlMaps property is available in versions that support XML mapping.
                // If the property is not present, this block will be skipped.
                var xmlMapsProperty = typeof(Workbook).GetProperty("XmlMaps");
                if (xmlMapsProperty != null)
                {
                    var xmlMaps = xmlMapsProperty.GetValue(workbook);
                    var addMethod = xmlMaps.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                    if (addMethod != null)
                    {
                        // Create the XML map
                        addMethod.Invoke(xmlMaps, new object[] { "MyCustomMap", schemaXml });
                    }
                }
            }

            // Save the workbook
            string outputPath = "MappedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
