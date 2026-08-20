// Title: Export an Excel XML map to a UTF‑8 file using Aspose.Cells C#
// Description: Loads a workbook, confirms the presence of an XML map, and uses Aspose.Cells ExportXml with a FileStream to write the map as UTF‑8. The sample reads the first three bytes of the output to verify the UTF‑8 BOM.
// Keywords: Aspose.Cells ExportXml UTF-8 | C# export XML map | Excel XML map to UTF-8 | ExportXml FileStream example | UTF-8 BOM verification Aspose
// Common Searches: Aspose.Cells export XML map UTF-8 | C# ExportXml with UTF-8 encoding | How to write UTF-8 BOM when exporting XML from Excel | Check encoding of exported XML Aspose.Cells
// Developer Intent: Generate a UTF‑8 encoded XML file from the first XML map in an Excel workbook.
// Use Cases: Provide UTF‑8 XML for APIs that require a BOM. | Create XML files for downstream data pipelines that expect UTF‑8 encoding. | Validate export encoding automatically after generating the file.
// AI Prompts: Write C# code that loads an Excel file, verifies XML maps, and exports the first map to a UTF‑8 XML file with Aspose.Cells. | Explain how Aspose.Cells determines the output encoding in ExportXml and how to guarantee a UTF‑8 BOM. | Show how to export a specific XML map by name and add error handling for missing maps.

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, confirms the presence of an XML map, and uses Aspose.Cells ExportXml with a FileStream to write the map as UTF‑8. The sample reads the first three bytes of the output to verify the UTF‑8 BOM.
class ExportXmlUtf8Demo
{
    static void Main()
    {
        // Path to the workbook that should contain an XML map
        string inputPath = "InputWithMap.xlsx";

        // Verify the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook wb;
        try
        {
            // Load the workbook
            wb = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Ensure there is at least one XML map in the workbook
        if (wb.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XML map found in the workbook.");
            return;
        }

        // Get the name of the first XML map
        string mapName = wb.Worksheets.XmlMaps[0].Name;

        // Define the output XML file path
        string outputPath = "ExportedUtf8.xml";

        try
        {
            // Export the XML using a FileStream (UTF‑8 encoding is used internally)
            using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                wb.ExportXml(mapName, fs);
            }

            // Verify that the file starts with a UTF‑8 BOM
            byte[] bom = new byte[3];
            using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
            {
                fs.Read(bom, 0, 3);
            }
            bool hasBom = bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF;
            Console.WriteLine($"Export completed. UTF‑8 BOM present: {hasBom}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during export: {ex.Message}");
        }
    }
}
