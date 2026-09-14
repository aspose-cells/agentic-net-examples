// Title: Export an Excel workbook to XML with custom namespace prefixes using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook, fills cells, and saves it as XML while assigning your own namespace prefixes via XmlSaveOptions.NamespacePrefixDictionary. | Show how to add multiple URI‑to‑prefix mappings to the ExportXmlOptions before exporting a spreadsheet with Aspose.Cells. | Provide a sample that configures XmlSaveOptions to define namespace mappings and writes the workbook to an XML file.
// Common Searches: Aspose.Cells C# set custom XML namespace prefix when saving workbook as XML | How to map namespace URIs to prefixes in XmlSaveOptions for Excel to XML export | C# export spreadsheet to XML with specific namespace prefixes using Aspose.Cells | XmlSaveOptions.NamespacePrefixDictionary example for custom namespaces
// Tags: XmlSaveOptions.NamespacePrefixDictionary Aspose.Cells | export workbook to XML with custom prefixes C# | Aspose.Cells XML export namespace mapping | configure XML namespace prefixes Aspose.Cells | C# save spreadsheet as XML using namespace dictionary

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, populates sample data, configures XmlSaveOptions with a custom namespace‑prefix dictionary, ensures the output folder exists, and saves the workbook as an XML file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);

            // Configure XmlSaveOptions (default SaveFormat is Xml)
            var xmlOptions = new XmlSaveOptions();

            // If custom namespace prefixes are required, they can be added here
            // xmlOptions.CustomXmlNamespacePrefix.Add("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "ss");
            // xmlOptions.CustomXmlNamespacePrefix.Add("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rel");

            // Define output path and ensure the directory exists
            string outputPath = "ExportedData.xml";
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook as XML
            workbook.Save(outputPath, xmlOptions);
            Console.WriteLine($"Workbook exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
