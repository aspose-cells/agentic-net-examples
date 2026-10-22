// Title: Assign a custom namespace prefix to an XML map in an Aspose.Cells workbook (C#)
// AI Prompts: Create an XML map with a specific namespace URI, set its Prefix property to a custom value, and save the workbook using Aspose.Cells for .NET. | Use reflection to retrieve the XmlMaps collection from a Workbook, add a named XML map, assign a custom prefix, and persist the file.
// Common Searches: C# Aspose.Cells how to set a prefix for an XML map namespace | adding an XML map with custom namespace prefix in Aspose.Cells workbook | reflection based access to XmlMaps collection Aspose.Cells example | save Excel file after configuring XML map namespace in .NET | example of setting XML map prefix for namespaced XML in Aspose.Cells
// Tags: set XML map prefix Aspose.Cells C# | add XML map with namespace URI Aspose.Cells | reflection access to Workbook.XmlMaps | configure XML map namespace in Excel workbook | save workbook after XML map setup Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample creates a Workbook, uses reflection to obtain the XmlMaps collection, adds an XML map with a defined namespace URI, assigns a custom Prefix (e.g., "inv"), and saves the workbook as CustomNamespacePrefix.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Attempt to add an XML map only if the XmlMaps property exists (some older versions may lack it)
            var xmlMapsProp = typeof(Workbook).GetProperty("XmlMaps");
            if (xmlMapsProp != null)
            {
                // Define the XML namespace URI that the XML map will use
                string xmlNamespace = "http://schemas.example.com/Invoice";

                // Get the XmlMapCollection instance via reflection
                var xmlMaps = xmlMapsProp.GetValue(workbook);
                dynamic maps = xmlMaps; // use dynamic to call Add and indexer at runtime

                // Add an XML map to the workbook with the specified namespace
                int mapIndex = maps.Add("InvoiceMap", xmlNamespace);

                // Retrieve the created XML map
                dynamic xmlMap = maps[mapIndex];

                // Set a custom namespace prefix (e.g., "inv") for the XML map
                xmlMap.Prefix = "inv";
            }

            // Save the workbook
            string outputPath = "CustomNamespacePrefix.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
