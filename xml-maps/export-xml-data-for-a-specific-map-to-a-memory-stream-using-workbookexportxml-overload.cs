// Title: Export a named XML map from an Excel workbook to a MemoryStream with Aspose.Cells (C#)
// AI Prompts: Write C# code that loads an .xlsx file, uses Workbook.ExportXml to export a specific XML map directly into a MemoryStream. | Show how to retrieve a named XML map from a workbook and return its XML content as a byte array using Aspose.Cells. | Provide a C# example that exports an XML map to a temporary file, copies it into a MemoryStream, and then deletes the temporary file.
// Common Searches: Aspose.Cells export specific XML map to stream in C# | How to get XML map bytes from an Excel file using Aspose.Cells | C# code for exporting a named XML map without saving a file | Workbook.ExportXml overload for named map to MemoryStream | Read exported XML map from temporary file using Aspose.Cells
// Tags: export named XML map to MemoryStream Aspose.Cells | Workbook.ExportXml overload for specific map | load workbook and extract XML map C# | temporary file handling after ExportXml | convert exported XML map to byte array Aspose.Cells | read XML map content from MemoryStream C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, uses Workbook.ExportXml to export a specified XML map to a temporary file, copies the file into a MemoryStream, optionally converts it to a byte array, displays the XML, and finally removes the temporary file.
class ExportMapToXml
{
    static void Main()
    {
        try
        {
            // Path to the input workbook.
            string inputPath = "input.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the map definition.
            Workbook workbook = new Workbook(inputPath);

            // Name of the map to export (must exist in the workbook).
            string mapName = "MyMap";

            // Temporary file to store the exported XML.
            string tempXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xml");

            try
            {
                // Export the specified map to the temporary XML file.
                workbook.ExportXml(tempXmlPath, mapName);

                // Read the exported XML into a memory stream.
                using (FileStream fileStream = new FileStream(tempXmlPath, FileMode.Open, FileAccess.Read))
                using (MemoryStream xmlStream = new MemoryStream())
                {
                    fileStream.CopyTo(xmlStream);
                    xmlStream.Position = 0;

                    // Optional: display the XML content.
                    using (StreamReader reader = new StreamReader(xmlStream))
                    {
                        string xmlContent = reader.ReadToEnd();
                        Console.WriteLine("Exported XML:");
                        Console.WriteLine(xmlContent);
                    }

                    // If raw bytes are needed:
                    // byte[] xmlBytes = xmlStream.ToArray();
                }
            }
            finally
            {
                // Clean up the temporary file.
                if (File.Exists(tempXmlPath))
                {
                    File.Delete(tempXmlPath);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
