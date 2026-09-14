// Title: Check required XML element XPaths have mapped cells before using ExportToXml with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook, iterates a list of required XPath strings, uses XmlMap.GetMappedCells to confirm each has at least one mapped cell, and only then calls ExportToXml. | Modify the validation loop to accumulate all missing XPath elements and output a single error report instead of terminating on the first missing mapping. | Create a version that processes every XML map in the workbook, validates required element mappings for each map, and exports each map to a separate XML file.
// Common Searches: Aspose.Cells how to ensure all required XML map elements are mapped before exporting to XML | C# validate XML map required XPaths with GetMappedCells in Aspose.Cells | Check missing cell mappings for XML elements in an Excel workbook using Aspose.Cells | Export workbook to XML only after confirming required XML elements have mapped cells in .NET
// Tags: Aspose.Cells XML map required element validation | GetMappedCells XPath verification Aspose.Cells | ExportToXml conditional export Aspose.Cells | C# XML map completeness check | multiple XML maps validation Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, verifies that an XML map exists, and checks a predefined list of required XPath elements using XmlMap.GetMappedCells to ensure each has at least one mapped cell. If any required element lacks a mapping, validation fails; otherwise, the workbook data is exported to an XML file with ExportToXml and the workbook is saved, with exception handling for robustness.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputXmlPath = "output.xml";
            const string outputXlsxPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Use dynamic to access XML map members (may not be available in older versions).
            dynamic wbDynamic = workbook;

            // Ensure at least one XML map is defined.
            if (wbDynamic.XmlMaps == null || wbDynamic.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML maps found in the workbook.");
                return;
            }

            // Use the first XML map.
            dynamic xmlMap = wbDynamic.XmlMaps[0];

            // List of required XML element XPaths that must have mapped cells.
            List<string> requiredElements = new List<string>
            {
                "/Root/Customer/Name",
                "/Root/Customer/Address/Street",
                "/Root/Order/OrderID"
            };

            // Validate that each required element has at least one mapped cell.
            foreach (string xpath in requiredElements)
            {
                // Get the cell areas mapped to the current XPath.
                CellArea[] mappedAreas = xmlMap.GetMappedCells(xpath);

                // If no mapping exists, report validation failure and stop.
                if (mappedAreas == null || mappedAreas.Length == 0)
                {
                    Console.WriteLine($"Validation failed: No cell mapped for required element '{xpath}'.");
                    return;
                }
            }

            Console.WriteLine("All required XML elements have corresponding mapped cells.");

            // Export the workbook data to an XML file using the validated map.
            xmlMap.ExportToXml(outputXmlPath);
            Console.WriteLine($"XML exported to '{outputXmlPath}'.");

            // Optionally save the workbook after export.
            workbook.Save(outputXlsxPath);
            Console.WriteLine($"Workbook saved as '{outputXlsxPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
