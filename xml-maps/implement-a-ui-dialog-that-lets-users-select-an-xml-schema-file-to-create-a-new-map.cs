// Title: Create a workbook and add an XML map from a user‑selected XSD file with Aspose.Cells for .NET
// Common Searches: C# console app prompt for XSD path and add XML map using Aspose.Cells | how to add an XML map to a new workbook from a file path with Aspose.Cells | Aspose.Cells save workbook as XLSX after adding XML map from schema | validate XSD file existence before creating XML map in C# | set friendly name for XML map based on XSD filename Aspose.Cells
// Tags: add xml map from xsd Aspose.Cells C# | create new workbook with xml map Aspose.Cells | assign xml map friendly name from schema filename Aspose.Cells | save workbook as xlsx after xml map addition Aspose.Cells | validate xsd file existence before adding xml map C#

using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapCreatorApp
{
    // A C# console program asks the user for an XSD file path, checks that the file exists, creates a new Workbook, adds an XML map using the selected schema, names the map after the XSD file, and saves the workbook as an XLSX file.
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                // Prompt the user to enter the full path of the XSD file.
                Console.WriteLine("Enter the full path of the XML Schema (.xsd) file:");
                string schemaPath = Console.ReadLine()?.Trim();

                // Validate the input.
                if (string.IsNullOrEmpty(schemaPath))
                {
                    Console.WriteLine("No file path was provided. Exiting.");
                    return;
                }

                // Ensure the file exists before proceeding.
                if (!File.Exists(schemaPath))
                {
                    Console.WriteLine("The specified file does not exist. Exiting.");
                    return;
                }

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Add the XML map using the selected schema file.
                // The Add method expects a path or URL to the XSD file.
                int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Optionally set a friendly name for the map.
                string mapName = Path.GetFileNameWithoutExtension(schemaPath);
                xmlMap.Name = mapName;

                // Save the workbook to demonstrate that the map was created.
                string outputPath = Path.Combine(Path.GetDirectoryName(schemaPath) ?? Environment.CurrentDirectory,
                                                "WorkbookWithXmlMap.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"XML Map \"{xmlMap.Name}\" created successfully.");
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message.
                Console.WriteLine($"An error occurred while creating the XML map: {ex.Message}");
            }
        }
    }
}
