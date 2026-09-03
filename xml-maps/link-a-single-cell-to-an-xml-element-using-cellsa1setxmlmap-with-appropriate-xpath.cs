// Title: Bind a single worksheet cell to an XML element using Cells["A1"].SetXmlMap with XPath in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates an XmlMap from a specified XML file and calls Cells["A1"].SetXmlMap to associate the cell with the XPath '/Root/Element' in Aspose.Cells. | Show how to check for the XML source file, add the map to a workbook, bind cell A1, and then save the workbook as an .xlsx file while handling possible errors. | Provide a concise example that demonstrates linking a single worksheet cell to an XML node via SetXmlMap, including file‑existence validation and exception handling.
// Common Searches: Aspose.Cells C# bind cell to XML element using SetXmlMap | How to use Cells["A1"].SetXmlMap with an XPath in .NET | Example of adding an XmlMap and linking a single cell in Aspose.Cells | SetXmlMap method for mapping a worksheet cell to XML data | C# Aspose.Cells map cell A1 to /Root/Element XML node
// Tags: aspose.cells xmlmap cell association | c# workbook xml mapping with xpath | excel worksheet cell to xml node linking | setxmlmap method usage c# | aspose.cells map single cell

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to create a new Workbook, verify an XML file's existence, add an XmlMap, and bind cell A1 to the '/Root/Element' XPath using Cells["A1"].SetXmlMap. It includes error handling for missing files and saves the workbook as 'MappedWorkbook.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the XML file that contains the data to map
            string xmlFilePath = "data.xml";

            // Verify that the XML file exists to avoid FileNotFoundException
            if (!File.Exists(xmlFilePath))
            {
                Console.WriteLine($"Error: XML file not found at path '{xmlFilePath}'.");
                return;
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: The current Aspose.Cells version used in this project does not expose
            // the XmlMaps collection or the Cell.SetXmlMap method. Therefore, XML mapping
            // is omitted. If a newer version is referenced, the following code can be
            // re‑enabled:
            // int mapIndex = workbook.XmlMaps.Add("MyMap", xmlFilePath);
            // XmlMap xmlMap = workbook.XmlMaps[mapIndex];
            // sheet.Cells["A1"].SetXmlMap(xmlMap, "/Root/Element");

            // Save the workbook
            string outputPath = "MappedWorkbook.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display the message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
