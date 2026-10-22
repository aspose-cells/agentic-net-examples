// Title: List all XML maps in an Excel workbook and print each map’s name and root element using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates through the workbook's XmlMapCollection, and writes each map’s Name and RootElementName to the console. | Show how to use dynamic binding in C# to safely access the XmlMaps property when the Aspose.Cells version does not expose it directly, including fallback error handling. | Create a robust example that checks for file existence, catches RuntimeBinderException for missing XmlMaps, and logs XML map details for debugging purposes.
// Common Searches: C# Aspose.Cells how to get list of XML maps from workbook | retrieve XmlMap name and root element using Aspose.Cells .NET | debug xml maps in Excel file with Aspose.Cells dynamic XmlMaps | enumerate XmlMapCollection when XmlMaps property not available in Aspose.Cells | list XML map definitions in an .xlsx using Aspose.Cells API
// Tags: enumerate xml maps Aspose.Cells | log xml map name root element C# | dynamic access XmlMaps Aspose.Cells | handle RuntimeBinderException Aspose.Cells | debug excel xml maps .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, uses dynamic binding to obtain the XmlMapCollection, iterates over each XML map, and writes the map's Name and RootElementName to the console. It includes checks for missing files and graceful handling of versions that lack the XmlMaps property.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Use dynamic to access XmlMaps (available in newer versions)
            dynamic dynWorkbook = workbook;

            try
            {
                // Retrieve the XML map collection
                var xmlMaps = dynWorkbook.XmlMaps as System.Collections.IEnumerable;

                if (xmlMaps != null)
                {
                    // Enumerate each XML map and output its details
                    foreach (var map in xmlMaps)
                    {
                        // Use dynamic again to access map properties safely
                        dynamic dynMap = map;
                        Console.WriteLine($"Map Name: {dynMap.Name}, Root Element: {dynMap.RootElementName}");
                    }
                }
                else
                {
                    Console.WriteLine("No XML maps are attached to this workbook.");
                }
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // XmlMaps property is not available in the referenced Aspose.Cells version
                Console.WriteLine("The XmlMaps feature is not supported by the current Aspose.Cells version.");
            }

            // Uncomment the line below to save any changes made to the workbook
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
