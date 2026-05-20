using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsXmlMapPreserveWhitespaceDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xml";
                const string outputPath = "output.xml";

                // Verify that the source XML file exists before attempting to load it.
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Enable XML mapping during load.
                var loadOptions = new XmlLoadOptions
                {
                    IsXmlMap = true
                };

                // Load the XML file into a workbook.
                var workbook = new Workbook(sourcePath, loadOptions);

                // Ensure at least one XML map is present.
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps were found in the loaded workbook.");
                    return;
                }

                // Use the first XML map for export.
                var xmlMap = workbook.Worksheets.XmlMaps[0];

                // Configure save options.
                var saveOptions = new XmlSaveOptions
                {
                    XmlMapName = xmlMap.Name
                };

                // Save the workbook as an XML file.
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine("XML exported with whitespace preserved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}