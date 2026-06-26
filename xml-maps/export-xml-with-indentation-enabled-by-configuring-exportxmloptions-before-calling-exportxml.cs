using System;
using System.IO;
using Aspose.Cells;
using System.Xml.Linq;

class ExportXmlWithIndentation
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_indented.xml";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load workbook containing an XML map
            Workbook workbook = new Workbook(inputPath);

            // Ensure at least one XML map is present
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XML map found in the workbook.");
                return;
            }

            // Use the first XML map's name for export
            string mapName = workbook.Worksheets.XmlMaps[0].Name;

            // Export XML (no direct indentation options in older versions)
            workbook.ExportXml(mapName, outputPath);

            // Re‑format the exported XML with indentation
            try
            {
                if (File.Exists(outputPath))
                {
                    XDocument doc = XDocument.Load(outputPath);
                    doc.Save(outputPath); // Saves with default indented formatting
                }
            }
            catch (Exception fmtEx)
            {
                Console.WriteLine($"XML formatting failed: {fmtEx.Message}");
            }

            Console.WriteLine("XML exported successfully with indentation.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}