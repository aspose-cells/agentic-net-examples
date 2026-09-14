// Title: Export the first XML map from an Excel workbook to a separate .xml file while preserving its original schema using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens a .xlsx file with Aspose.Cells, verifies the presence of XmlMaps, and calls ExportData on the first map with the keepSchema flag set to true. | Create a method that uses reflection to obtain the XmlMaps collection from a Workbook object and saves the map data to a specified XML file while retaining the original schema. | Develop a console application that checks for an input workbook, handles cases where no XML maps exist, and exports the first XML map to an output .xml file with schema preservation.
// Common Searches: aspnet c# export xml map from excel workbook preserving schema Aspose.Cells | how to use Aspose.Cells ExportData to keep original XML schema | C# code sample for exporting first XML map in .xlsx to .xml file | Aspose.Cells XmlMaps collection reflection example | export xml map data to external file with schema using Aspose.Cells .NET
// Tags: asp.net xml map handling Aspose.Cells | dynamic retrieval of XmlMaps collection | save xml map data to external .xml file | retain xml schema layout during export | initial xml map extraction from workbook

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// // Loads an Excel workbook, uses reflection to access its XmlMaps collection, checks for available maps, and exports the first XML map's data to a separate .xml file while preserving the original schema structure.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "exportedData.xml";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        try
        {
            // Load the workbook that may contain XML maps
            Workbook workbook = new Workbook(inputPath);

            // Use reflection to obtain the XmlMaps collection (API may vary by version)
            PropertyInfo xmlMapsProp = workbook.GetType().GetProperty("XmlMaps", BindingFlags.Public | BindingFlags.Instance);
            if (xmlMapsProp == null)
            {
                Console.WriteLine("The current Aspose.Cells version does not support XML maps.");
                return;
            }

            dynamic xmlMaps = xmlMapsProp.GetValue(workbook);
            if (xmlMaps == null || xmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps are present in the workbook.");
                return;
            }

            // Export the first XML map's data to an external XML file.
            // The second argument (true) preserves the original schema structure.
            xmlMaps[0].ExportData(outputPath, true);
            Console.WriteLine($"XML data exported successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
