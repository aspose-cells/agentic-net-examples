// Title: Re‑add an existing XML map with a new XSD schema in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that locates an XML map by name (ignoring case), deletes it, and adds a replacement map using a specified XSD file with Aspose.Cells. | Show how to use reflection to obtain the XmlMaps collection, handle missing workbook or XSD files, and save the workbook after the map update. | Provide a robust example that verifies the added XML map’s schema file, creates the output directory if needed, and logs success or error messages.
// Common Searches: aspnet remove xml map and add new xsd schema in existing workbook using Aspose.Cells | c# update xml map definition in Excel file without recreating workbook | how to replace xml map in Aspose.Cells when map name case differs | using reflection to access XmlMaps collection Aspose.Cells .NET
// Tags: delete existing xml map Aspose.Cells | insert xml map from xsd Aspose.Cells | update xml map schema .NET | reflective access to XmlMaps collection | case‑insensitive xml map lookup

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, validates the workbook and XSD file paths, uses reflection to retrieve the XmlMaps collection, removes any map whose name matches the supplied name (case‑insensitive), adds a new XML map with the provided XSD schema, logs the schema information of the added map, ensures the output directory exists, saves the updated workbook, and reports errors with clear messages.
class UpdateXmlMap
{
    static void Main()
    {
        // Paths (replace with actual file locations)
        string inputFile = @"{inputFile}";   // Existing workbook
        string outputFile = @"{outputFile}"; // Workbook to save after update
        string xsdPath = @"{xsdPath}";       // New XSD schema file
        string mapName = "{mapName}";        // Name of the XML map to update

        try
        {
            // Verify input files exist
            if (!File.Exists(inputFile))
                throw new FileNotFoundException("Input workbook not found.", inputFile);
            if (!File.Exists(xsdPath))
                throw new FileNotFoundException("XSD schema file not found.", xsdPath);

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Access XmlMaps via dynamic to stay compatible with different library versions
            var xmlMapsProp = workbook.GetType().GetProperty("XmlMaps");
            if (xmlMapsProp == null)
                throw new NotSupportedException("The loaded Aspose.Cells version does not support XML maps.");

            dynamic xmlMaps = xmlMapsProp.GetValue(workbook);
            if (xmlMaps == null)
                throw new InvalidOperationException("Failed to retrieve XmlMaps collection.");

            // Find and remove any existing map with the same name (case‑insensitive)
            int mapIndex = -1;
            for (int i = 0; i < (int)xmlMaps.Count; i++)
            {
                dynamic map = xmlMaps[i];
                if (string.Equals((string)map.Name, mapName, StringComparison.OrdinalIgnoreCase))
                {
                    mapIndex = i;
                    break;
                }
            }
            if (mapIndex != -1)
                xmlMaps.RemoveAt(mapIndex);

            // Add the new XML map using the provided XSD schema
            int newMapIndex = xmlMaps.Add(mapName, xsdPath);

            // Optional verification
            dynamic newMap = xmlMaps[newMapIndex];
            string schemaInfo = newMap.GetType().GetProperty("SchemaFile") != null
                ? (string)newMap.SchemaFile
                : (newMap.GetType().GetProperty("SchemaFileName") != null
                    ? (string)newMap.SchemaFileName
                    : "unknown");

            Console.WriteLine($"XML map '{newMap.Name}' added with schema '{schemaInfo}'.");

            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(outputFile);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the workbook with the updated XML map
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
