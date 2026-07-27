using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapRemoval
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "InputWithXmlMaps.xlsx";
            const string outputPath = "OutputAfterRemoval.xlsx";
            const string mapNameToRemove = "Employee";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Attempt to remove the specified XML map
                bool removed = RemoveXmlMapByName(workbook, mapNameToRemove);

                Console.WriteLine(removed
                    ? $"XML map \"{mapNameToRemove}\" was removed."
                    : $"XML map \"{mapNameToRemove}\" not found.");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <summary>
        /// Searches the workbook's XmlMap collection for a map with the specified name
        /// and removes it if found.
        /// </summary>
        /// <param name="workbook">The workbook containing the XmlMaps.</param>
        /// <param name="mapName">The name of the XmlMap to remove.</param>
        /// <returns>True if the map was found and removed; otherwise false.</returns>
        static bool RemoveXmlMapByName(Workbook workbook, string mapName)
        {
            // Access the XmlMapCollection from the workbook
            XmlMapCollection xmlMaps = workbook.Worksheets.XmlMaps;

            // Iterate through the collection to locate the map by name
            for (int i = 0; i < xmlMaps.Count; i++)
            {
                XmlMap map = xmlMaps[i];
                if (string.Equals(map.Name, mapName, StringComparison.OrdinalIgnoreCase))
                {
                    // Remove the map at the found index
                    xmlMaps.RemoveAt(i);
                    return true;
                }
            }

            // Map with the specified name was not found
            return false;
        }
    }
}