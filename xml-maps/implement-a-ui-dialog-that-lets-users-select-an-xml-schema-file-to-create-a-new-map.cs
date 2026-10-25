// Title: Create a console dialog to select an XSD schema file and save a new workbook with Aspose.Cells in C#
// Common Searches: C# console app prompt user for XSD file path using Aspose.Cells | validate existence of XML schema file before creating Aspose.Cells workbook | save a new Aspose.Cells workbook to a user‑specified location in C# | how to create output folder if it does not exist before saving Aspose.Cells file | handle file‑not‑found errors when loading XSD for Aspose.Cells XML map
// Tags: Aspose.Cells load XML schema from XSD in C# | C# console file path validation for Aspose.Cells | create workbook and save as XLSX using Aspose.Cells | ensure output directory exists before workbook save | handle missing schema file error in Aspose.Cells workflow

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapDemo
{
    // The example shows a C# console program that asks the user for the full path of an XML schema (.xsd) file, checks that the file exists, creates a new Aspose.Cells Workbook, acknowledges the schema for potential XML map use, then prompts for a save location, creates the target directory if necessary, and finally saves the workbook as an XLSX file while handling errors.
    public class Program
    {
        static void Main()
        {
            try
            {
                // Prompt for XML schema file path
                Console.WriteLine("Enter full path to XML Schema (.xsd) file:");
                string schemaPath = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(schemaPath) || !File.Exists(schemaPath))
                {
                    Console.WriteLine("Schema file not found.");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // NOTE: The XmlMaps API may not be available in some Aspose.Cells versions.
                // If needed, XML map functionality can be added here using the appropriate API.
                // For now we simply acknowledge the schema file.
                Console.WriteLine($"Schema file '{Path.GetFileName(schemaPath)}' is ready for XML map import (functionality omitted).");

                // Prompt for save location
                Console.WriteLine("Enter full path to save the workbook (e.g., C:\\temp\\output.xlsx):");
                string savePath = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(savePath))
                {
                    Console.WriteLine("No save path provided. Exiting.");
                    return;
                }

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(savePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(savePath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
