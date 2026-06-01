using System;
using System.IO;
using Aspose.Cells;

namespace BatchXmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the source Excel workbooks
            string sourceDirectory = @"C:\InputWorkbooks";

            // Directory where the exported XML files will be saved
            string outputDirectory = @"C:\ExportedXml";

            // Path to the XML schema (XSD) that defines the XML map to be applied
            string xmlMapPath = @"C:\Schema\SampleMap.xsd";

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Get all Excel files in the source directory (you can adjust the pattern as needed)
            string[] workbookFiles = Directory.GetFiles(sourceDirectory, "*.xlsx");

            foreach (string workbookFile in workbookFiles)
            {
                try
                {
                    // Load the workbook from file
                    Workbook workbook = new Workbook(workbookFile);

                    // If the workbook does not contain any XML maps, add the required map
                    if (workbook.Worksheets.XmlMaps.Count == 0)
                    {
                        // Add the XML map using the XSD file path
                        workbook.Worksheets.XmlMaps.Add(xmlMapPath);
                    }

                    // Retrieve the first (or the intended) XML map
                    XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                    // Determine the output XML file path (same name as workbook, different extension)
                    string outputXmlPath = Path.Combine(
                        outputDirectory,
                        Path.GetFileNameWithoutExtension(workbookFile) + ".xml");

                    // Export the XML data using the map's name
                    workbook.ExportXml(xmlMap.Name, outputXmlPath);

                    Console.WriteLine($"Exported XML for '{Path.GetFileName(workbookFile)}' to '{outputXmlPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(workbookFile)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch XML export completed.");
        }
    }
}