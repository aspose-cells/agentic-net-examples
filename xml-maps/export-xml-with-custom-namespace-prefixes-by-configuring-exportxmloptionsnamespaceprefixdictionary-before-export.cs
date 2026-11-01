// Title: Export an Excel workbook to XML with custom namespace prefixes using Aspose.Cells ExportXmlOptions in C#
// AI Prompts: Write C# code that adds custom prefix‑to‑namespace entries to ExportXmlOptions.NamespacePrefixDictionary and then saves the workbook as XML with Aspose.Cells. | Update the provided Aspose.Cells example to define a namespace prefix (e.g., "ns" for "http://example.com/schema") before exporting the workbook to XML. | Generate a complete C# snippet that creates a workbook, fills it with data, configures ExportXmlOptions with custom namespace mappings, and writes the XML file to disk.
// Common Searches: Aspose.Cells C# export to XML with custom namespace prefix mapping | How to set NamespacePrefixDictionary in ExportXmlOptions for XML export | C# example for adding XML namespace prefixes when exporting Excel with Aspose.Cells | Export Excel workbook to XML using Aspose.Cells and specify namespace prefixes
// Tags: Aspose.Cells custom XML namespace prefixes | ExportXmlOptions custom namespace mapping C# | C# export workbook to XML with Aspose.Cells | XML namespace dictionary Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

namespace ExportXmlWithCustomNamespaces
{
    // The sample creates a new workbook, populates it with sample data, prepares ExportXmlOptions (or XmlSaveOptions), ensures the output directory exists, and saves the workbook as an XML file; you can inject custom namespace prefix mappings via ExportXmlOptions.NamespacePrefixDictionary before the save operation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and set its name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Name");
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue("Alice");
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue("Bob");

                // Configure XML save options
                XmlSaveOptions xmlOptions = new XmlSaveOptions();

                // Define output path
                string outputPath = "CustomNamespaceExport.xml";

                // Ensure the output directory exists (if any)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to XML using the configured options
                workbook.Save(outputPath, xmlOptions);

                Console.WriteLine($"Workbook exported to XML at: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
