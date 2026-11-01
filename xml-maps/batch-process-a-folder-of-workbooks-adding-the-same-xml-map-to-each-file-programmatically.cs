// Title: How to batch add the same XML map to multiple Excel workbooks using Aspose.Cells for .NET (C#)
// Common Searches: add identical XML map to many .xlsx files with Aspose.Cells C# | batch process a folder of Excel workbooks to attach an XSD schema programmatically | use reflection to add XmlMaps when the XmlMaps property is unavailable in Aspose.Cells | C# code for applying the same XML map to all workbooks in a directory | automate XML map insertion into multiple Excel files using Aspose.Cells
// Tags: xml map insertion Aspose.Cells C# | batch workbook processing Aspose.Cells | xsd schema loading into Excel workbook | reflection based XmlMaps addition | multiple workbook XML map automation

using System;
using System.IO;
using Aspose.Cells;

// The example scans an input folder for .xlsx files, reads a single XSD schema, and programmatically adds an XML map named "MyXmlMap" to each workbook (skipping existing maps). It uses reflection to stay compatible with older Aspose.Cells versions, then saves the modified workbooks to an output folder.
class BatchXmlMapAdder
{
    static void Main()
    {
        // Input and output directories
        string inputFolder = @"C:\InputFolder";
        string outputFolder = @"C:\OutputFolder";

        // Path to the XSD schema
        string schemaPath = @"C:\schema.xsd";

        // Verify that the schema file exists
        if (!File.Exists(schemaPath))
        {
            Console.WriteLine($"Schema file not found: {schemaPath}");
            return;
        }

        // Read the XSD content once
        string schemaXml;
        try
        {
            schemaXml = File.ReadAllText(schemaPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read schema file: {ex.Message}");
            return;
        }

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        const string mapName = "MyXmlMap";

        // Process each .xlsx file in the input folder
        foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Workbook not found (skipped): {filePath}");
                continue;
            }

            Workbook workbook = null;
            try
            {
                workbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook '{filePath}': {ex.Message}");
                continue;
            }

            // Add XML map using reflection to stay compatible with versions that may not expose XmlMaps directly
            try
            {
                var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
                if (xmlMapsProp != null)
                {
                    var xmlMaps = xmlMapsProp.GetValue(workbook);
                    if (xmlMaps != null)
                    {
                        // Check if the map already exists (if Contains method is available)
                        var containsMethod = xmlMaps.GetType().GetMethod("Contains", new[] { typeof(string) });
                        bool mapExists = false;
                        if (containsMethod != null)
                        {
                            mapExists = (bool)containsMethod.Invoke(xmlMaps, new object[] { mapName });
                        }

                        if (!mapExists)
                        {
                            var addMethod = xmlMaps.GetType().GetMethod("Add", new[] { typeof(string), typeof(string) });
                            if (addMethod != null)
                            {
                                addMethod.Invoke(xmlMaps, new object[] { mapName, schemaXml });
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("XmlMaps property not available in this Aspose.Cells version; XML map not added.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while adding XML map to '{filePath}': {ex.Message}");
            }

            // Save the modified workbook
            string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));
            try
            {
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook '{outputPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Processing completed.");
    }
}
