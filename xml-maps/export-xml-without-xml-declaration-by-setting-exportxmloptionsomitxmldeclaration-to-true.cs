using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class ExportXmlWithoutDeclaration
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleData";

                sheet.Cells["A1"].PutValue("Id");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Add a simple XML map to the workbook
                int mapIndex = workbook.Worksheets.XmlMaps.Add("<Schema><Element><Id/><Name/></Element></Schema>");
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[mapIndex];
                xmlMap.Name = "SampleMap";

                // Define output file path
                string outputPath = "output_without_declaration.xml";

                // Ensure the output directory exists (handle possible null)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export XML (default includes declaration)
                workbook.ExportXml(xmlMap.Name, outputPath);

                // Remove XML declaration if present
                string xmlContent = File.ReadAllText(outputPath);
                if (xmlContent.StartsWith("<?xml", StringComparison.Ordinal))
                {
                    int declEnd = xmlContent.IndexOf("?>", StringComparison.Ordinal);
                    if (declEnd > -1)
                    {
                        xmlContent = xmlContent.Substring(declEnd + 2).TrimStart('\r', '\n');
                        File.WriteAllText(outputPath, xmlContent);
                    }
                }

                Console.WriteLine($"XML exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}