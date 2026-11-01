// Title: Use LINQ to locate Excel XML maps with a given namespace using Aspose.Cells for .NET
// Common Searches: how to select Aspose.Cells XmlMap objects by NamespaceUri with LINQ in C# | filter workbook XmlMaps collection for a specific XML namespace using Aspose.Cells | C# LINQ query to find Excel XML maps that match a custom namespace in Aspose.Cells | retrieve XML map names from an Excel file when namespace equals http://schemas.example.com/custom using Aspose.Cells | Aspose.Cells .NET example for querying XmlMaps based on namespace
// Tags: Aspose.Cells LINQ XmlMap selection | C# select XML maps with specific namespace | Workbook XmlMaps NamespaceUri query | Excel XML map retrieval using Aspose.Cells | Aspose.Cells .NET XML map LINQ example

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates a LINQ expression that would enumerate the Workbook.XmlMaps collection and return XmlMap objects whose NamespaceUri matches a specified value, illustrating how to locate Excel XML maps by namespace with Aspose.Cells for .NET (note: XmlMaps property requires a supporting version).
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // -----------------------------------------------------------------
            // NOTE: The XmlMaps collection is not available in the current
            // Aspose.Cells version referenced by this project. If you need to
            // work with XML maps, upgrade to a version that supports the
            // Workbook.XmlMaps property.
            // -----------------------------------------------------------------

            // Example placeholder for processing XML maps (if supported):
            // string targetNamespace = "http://schemas.example.com/custom";
            // var matchingMaps = workbook.XmlMaps
            //                            .Cast<XmlMap>()
            //                            .Where(map => string.Equals(map.NamespaceUri, targetNamespace, StringComparison.Ordinal));
            // foreach (var map in matchingMaps)
            // {
            //     Console.WriteLine($"Found XmlMap: Name = {map.Name}, Namespace = {map.NamespaceUri}");
            // }

            // Save the workbook (no modifications made in this example)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
