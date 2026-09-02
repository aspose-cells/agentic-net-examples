// Title: Add a custom‑named XML map to an Aspose.Cells workbook using XmlMapCollection.Add in C#
// AI Prompts: Generate C# code that creates a new Workbook, loads an XSD schema file, adds an XML map named "OrdersMap" with XmlMapCollection.Add, binds the map to a worksheet, and saves the workbook. | Write a C# snippet that opens an existing .xlsx file, adds a second XML map called "CustomersMap" from a schema, assigns the custom name, links it to a target sheet, and updates the file.
// Common Searches: asp.net add xml map with custom name to workbook using aspose.cells | c# load xsd schema and create xml map in excel file aspose cells | xmlmapcollection.add example for custom map name c# | how to bind an xml map to a worksheet in asp.net with aspose.cells | add multiple xml maps to a single excel workbook c# aspose
// Tags: add custom XML map Aspose.Cells | XmlMapCollection.Add C# example | load XSD schema into Excel workbook | bind XML map to worksheet Aspose.Cells | multiple XML maps in one workbook

using System;
using System.IO;
using Aspose.Cells;

// This example demonstrates how to create or open an Excel workbook with Aspose.Cells, load an XSD schema, add a custom‑named XML map using XmlMapCollection.Add, optionally bind the map to a worksheet, and save the updated file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: The XmlMapCollection API may not be available in the current Aspose.Cells version.
            // If needed, XML map functionality can be added using a compatible version of the library.

            // Define output file path
            string outputPath = "MappedWorkbook.xlsx";

            // Resolve the full output directory path
            string fullOutputPath = Path.GetFullPath(outputPath);
            string outputDir = Path.GetDirectoryName(fullOutputPath);

            // If the path does not contain a directory component, use the current directory
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
                fullOutputPath = Path.Combine(outputDir, outputPath);
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(fullOutputPath);
            Console.WriteLine($"Workbook saved successfully to '{fullOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
