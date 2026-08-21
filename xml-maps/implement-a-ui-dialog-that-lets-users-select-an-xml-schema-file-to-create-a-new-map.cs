// Title: Add an XML Schema (XSD) as an XmlMap to a New Workbook via Console UI – Aspose.Cells for .NET
// Description: A C# console program that creates a Workbook, prompts the user for an XSD or XML schema file, validates the path, adds the schema as an XmlMap with a friendly name, and saves the workbook as WorkbookWithXmlMap.xlsx while handling errors.
// Keywords: Aspose.Cells XmlMap | C# add XSD to workbook | XML schema map Excel | console UI select XSD | save workbook with XmlMap | Aspose.Cells .NET example | XmlMaps.Add usage
// Common Searches: how to add an XSD as an XmlMap using Aspose.Cells | C# console app to import XML schema into Excel | Aspose.Cells XmlMaps.Add example code | create Excel workbook with XML map from file path | validate schema file before adding to Aspose.Cells workbook
// Developer Intent: Prompt the user for an XML schema file, add it as an XmlMap to a new workbook, and save the workbook.
// Use Cases: Enable end‑users to link custom XML schemas to Excel templates at runtime. | Automate generation of workbooks that require predefined XML‑Excel mappings for data exchange. | Build lightweight console utilities for batch creation of Excel files with embedded XmlMaps.
// AI Prompts: Generate a WinForms file‑picker dialog that lets users browse for an XSD and adds it as an XmlMap with Aspose.Cells. | Modify the console program to accept multiple schema paths and create a separate XmlMap for each before saving. | Provide detailed exception handling that distinguishes file‑not‑found, invalid XSD format, and Aspose.Cells mapping errors.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDialog
{
    // A C# console program that creates a Workbook, prompts the user for an XSD or XML schema file, validates the path, adds the schema as an XmlMap with a friendly name, and saves the workbook as WorkbookWithXmlMap.xlsx while handling errors.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (document creation)
                Workbook workbook = new Workbook();

                // Prompt user to enter the path of the XML schema (XSD) file
                Console.WriteLine("Enter the full path to the XML Schema file (XSD or XML):");
                string schemaPath = Console.ReadLine()?.Trim();

                // Validate the input
                if (string.IsNullOrEmpty(schemaPath) || !File.Exists(schemaPath))
                {
                    Console.WriteLine("The specified schema file does not exist. Exiting.");
                    return;
                }

                // Add the selected schema as an XmlMap to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Optionally set a friendly name for the map (using file name without extension)
                xmlMap.Name = Path.GetFileNameWithoutExtension(schemaPath);

                Console.WriteLine($"XML schema added as map '{xmlMap.Name}' (Index: {mapIndex}).");

                // Save the workbook to a file (document saving)
                string outputPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithXmlMap.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
