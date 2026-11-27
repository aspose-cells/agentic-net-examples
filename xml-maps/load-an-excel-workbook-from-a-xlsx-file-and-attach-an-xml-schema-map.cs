// Title: Programmatically load an .xlsx workbook and attach an XSD schema map with Aspose.Cells for .NET (C#)
// Common Searches: how to add an XML schema (XSD) to an existing Excel workbook using Aspose.Cells in C# | C# code to load a .xlsx file and attach an XmlMap from an XSD file with Aspose.Cells | Aspose.Cells XmlMaps.Add example for attaching XSD to a workbook and saving it | assign friendly name to XML map when adding schema to Excel workbook using Aspose.Cells | save workbook after attaching XML schema map with Aspose.Cells .NET
// Tags: load workbook and add XML schema map Aspose.Cells | XmlMaps.Add XSD attachment C# | assign name to XmlMap from schema file | save workbook with attached XML map Aspose.Cells | display XML map details Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapAttachmentDemo
{
    // Loads an existing .xlsx workbook, adds an XML schema map from an .xsd file using Workbook.Worksheets.XmlMaps.Add, optionally sets a friendly map name, prints map details, and saves the workbook with the attached XML map.
    class Program
    {
        static void Main(string[] args)
        {
            // Validate command‑line arguments
            if (args.Length != 3)
            {
                Console.WriteLine("Incorrect number of arguments.");
                Console.WriteLine("Usage: XmlMapAttachmentDemo.exe <inputWorkbook.xlsx> <schemaFile.xsd> <outputWorkbook.xlsx>");
                return;
            }

            string inputWorkbookPath = args[0];
            string schemaPath = args[1];
            string outputWorkbookPath = args[2];

            // Verify that the input files exist
            if (!File.Exists(inputWorkbookPath))
            {
                Console.WriteLine($"Error: Input workbook not found at \"{inputWorkbookPath}\"");
                return;
            }

            if (!File.Exists(schemaPath))
            {
                Console.WriteLine($"Error: XML schema file not found at \"{schemaPath}\"");
                return;
            }

            try
            {
                // Load the existing workbook
                Console.WriteLine($"Loading workbook \"{inputWorkbookPath}\"...");
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Add the XML schema map to the workbook's XmlMaps collection
                Console.WriteLine($"Adding XML schema map from \"{schemaPath}\"...");
                int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Optionally set a friendly name for the map (if not already set)
                if (string.IsNullOrEmpty(xmlMap.Name))
                {
                    string mapName = Path.GetFileNameWithoutExtension(schemaPath);
                    xmlMap.Name = mapName;
                    Console.WriteLine($"Assigned map name: \"{xmlMap.Name}\"");
                }
                else
                {
                    Console.WriteLine($"Map name already defined: \"{xmlMap.Name}\"");
                }

                // Display some information about the attached map
                Console.WriteLine("XML Map Details:");
                Console.WriteLine($"  Index          : {mapIndex}");
                Console.WriteLine($"  Name           : {xmlMap.Name}");
                Console.WriteLine($"  RootElement    : {xmlMap.RootElementName}");
                Console.WriteLine($"  DataBinding    : {(xmlMap.DataBinding != null ? "Exists" : "None")}");

                // Save the workbook with the attached XML map
                Console.WriteLine($"Saving workbook with attached XML map to \"{outputWorkbookPath}\"...");
                workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);

                Console.WriteLine("Operation completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
