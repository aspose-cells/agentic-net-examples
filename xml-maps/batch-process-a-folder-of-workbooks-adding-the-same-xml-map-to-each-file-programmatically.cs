using System;
using System.IO;
using Aspose.Cells;

namespace BatchXmlMapAdder
{
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the Excel workbooks to process
            string inputFolder = @"C:\InputWorkbooks";

            // Path to the XML schema (XSD) that defines the XML map
            string xmlSchemaPath = @"C:\Schema\sample.xsd";

            // Optional: output folder (can be same as input to overwrite)
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Verify that the XML schema file exists
            if (!File.Exists(xmlSchemaPath))
            {
                Console.WriteLine($"Schema file not found: {xmlSchemaPath}");
                return;
            }

            string xmlSchemaContent;
            try
            {
                // Read the XML schema content once (can be a file path or raw schema string)
                xmlSchemaContent = File.ReadAllText(xmlSchemaPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read schema file: {ex.Message}");
                return;
            }

            // Process each .xlsx file in the input folder
            foreach (string filePath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                try
                {
                    // Verify that the workbook file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"Workbook file not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Add the XML map to the workbook
                    int mapIndex = workbook.Worksheets.XmlMaps.Add(xmlSchemaContent);

                    // Set a friendly name for the map
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                    xmlMap.Name = Path.GetFileNameWithoutExtension(filePath) + "_Map";

                    // Determine the output file path (overwrite or new location)
                    string outputPath = Path.Combine(outputFolder, Path.GetFileName(filePath));

                    // Save the modified workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}