using System;
using System.IO;
using Aspose.Cells;

namespace XmlMapCreator
{
    // Main entry point
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Console.WriteLine("Enter full path to XML Schema file (XSD or XML):");
            string inputPath = Console.ReadLine()?.Trim('\"', ' ', '\t');

            if (string.IsNullOrWhiteSpace(inputPath))
            {
                Console.WriteLine("No path provided. Exiting.");
                return;
            }

            // Prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add the selected XML schema as a map
                int mapIndex = workbook.Worksheets.XmlMaps.Add(inputPath);
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];

                // Set a friendly name based on file name
                xmlMap.Name = Path.GetFileNameWithoutExtension(inputPath);

                // Determine output path
                string outputPath = Path.Combine(
                    Path.GetDirectoryName(inputPath) ?? Environment.CurrentDirectory,
                    $"{xmlMap.Name}_Mapped.xlsx");

                // Save the workbook
                workbook.Save(outputPath);

                Console.WriteLine($"XML map added and workbook saved to:\n{outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}