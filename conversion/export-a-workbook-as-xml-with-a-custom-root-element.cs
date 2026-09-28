// Title: How to export an Aspose.Cells workbook to an XML file in C# using XmlSaveOptions
// AI Prompts: Create C# code that builds an Aspose.Cells workbook, populates cells, and saves it as an XML document with XmlSaveOptions. | Write a C# routine that checks for the existence of the target folder and creates it if missing before exporting a workbook to XML using Aspose.Cells. | Generate a C# example that configures XmlSaveOptions for default XML export and saves the workbook to a specified path.
// Common Searches: Aspose.Cells C# export workbook to XML file example | How to use XmlSaveOptions for XML output in Aspose.Cells .NET | C# save Excel workbook as XML with Aspose.Cells and ensure output folder exists | Set custom root element when exporting Excel to XML using Aspose.Cells C# | XmlSaveOptions default settings for XML export in Aspose.Cells C#
// Tags: Aspose.Cells export workbook to XML | XmlSaveOptions configure XML output | C# ensure output directory before saving | Aspose.Cells create workbook programmatically | Export Excel data as XML using .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportXml
{
    // The example demonstrates creating an Aspose.Cells workbook, filling it with sample data, configuring XmlSaveOptions, ensuring the destination directory exists, and saving the workbook as an XML file while handling potential errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Age");
                sheet.Cells["A2"].PutValue("John");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Jane");
                sheet.Cells["B3"].PutValue(25);

                // Configure XML save options (SaveFormat is Xml by default, no need to set)
                XmlSaveOptions xmlOptions = new XmlSaveOptions();

                // Define output file path
                string outputPath = "ExportedWorkbook.xml";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to XML
                workbook.Save(outputPath, xmlOptions);

                Console.WriteLine($"Workbook successfully exported to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
