// Title: Export an Excel workbook to XML without the XML declaration using Aspose.Cells for .NET
// AI Prompts: Provide a C# snippet that disables the XML declaration when saving a workbook as XML using Aspose.Cells. | Show how to create the output folder and export the workbook to an XML file while preventing the <?xml?> header from being written.
// Common Searches: Aspose.Cells C# export Excel as XML without the <?xml?> version tag | How to turn off the XML declaration in Aspose.Cells XmlSaveOptions | Example of saving a workbook to XML with no header using Aspose.Cells .NET | C# code to generate XML from Excel without the XML prolog via Aspose.Cells | Remove XML prolog when converting Excel to XML with Aspose.Cells library
// Tags: Aspose.Cells XmlSaveOptions suppress XML header | export Excel workbook to XML without declaration C# | C# Aspose.Cells save workbook as XML without <?xml?> | OmitXmlDeclaration property Aspose.Cells .NET | XML export options for Excel using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, fills sample data, configures XmlSaveOptions with OmitXmlDeclaration = true, ensures the output directory exists, and saves the workbook to an XML file that omits the XML declaration.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);

            // Configure XML save options (default behavior)
            XmlSaveOptions saveOptions = new XmlSaveOptions();

            // Define output file path
            string outputPath = "ExportedData.xml";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Export the workbook to an XML file
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
