using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapCreator
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prompt user to enter the path of an XML schema (XSD) file
            string schemaPath = PromptForSchemaFile();
            if (string.IsNullOrEmpty(schemaPath) || !File.Exists(schemaPath))
            {
                Console.WriteLine("Invalid schema file path. Exiting.");
                return;
            }

            // Add the selected schema as an XML map to the workbook
            int mapIndex = workbook.Worksheets.XmlMaps.Add(schemaPath);
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

            // Set a friendly name for the map
            xmlMap.Name = Path.GetFileNameWithoutExtension(schemaPath);

            // Prompt user for a location to save the workbook
            string savePath = PromptForSaveFile();
            if (string.IsNullOrEmpty(savePath))
            {
                Console.WriteLine("No save location provided. Exiting.");
                return;
            }

            // Ensure the directory exists
            string dir = Path.GetDirectoryName(savePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            // Save the workbook with the new XML map
            workbook.Save(savePath);
            Console.WriteLine($"Workbook saved successfully to: {savePath}");
        }

        // Reads a file path from the console for the XML schema
        private static string PromptForSchemaFile()
        {
            Console.Write("Enter full path to the XML Schema (XSD) file: ");
            return Console.ReadLine()?.Trim();
        }

        // Reads a file path from the console for saving the workbook
        private static string PromptForSaveFile()
        {
            Console.Write("Enter full path where the workbook should be saved (e.g., C:\\Temp\\WorkbookWithXmlMap.xlsx): ");
            return Console.ReadLine()?.Trim();
        }
    }
}