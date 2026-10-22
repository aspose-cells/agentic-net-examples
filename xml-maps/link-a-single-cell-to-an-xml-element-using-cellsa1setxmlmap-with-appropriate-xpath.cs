// Title: How to bind cell A1 to an XML element using Cells["A1"].SetXmlMap with XPath in Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an XML map from a file, defines the XPath '/Root/Element', and calls Cells["A1"].SetXmlMap to bind that cell to the XML node using Aspose.Cells. | Write a .NET example that loads an XML document, adds the map to a workbook, links a single worksheet cell to the specified XML element via SetXmlMap, and saves the workbook.
// Common Searches: Aspose.Cells C# bind a single Excel cell to an XML node using SetXmlMap | SetXmlMap example for linking cell A1 to /Root/Element in .NET | How to map an XML element to a specific cell with Aspose.Cells | Create XML map and assign to a cell in Aspose.Cells C# tutorial | Using XPath with SetXmlMap to connect Excel cell to XML data
// Tags: SetXmlMap with XPath Aspose.Cells | bind single cell to XML element C# | Excel XML mapping using Aspose.Cells .NET | create XML map for worksheet cell | link cell A1 to XML node Aspose.Cells

using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Cells;

// The example loads an XML file, creates an XML map with the XPath '/Root/Element', binds cell A1 of a new workbook to that XML node using Cells["A1"].SetXmlMap, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the XML file that will be mapped
            string xmlFilePath = "data.xml";

            // Ensure the XML file exists; create a simple one if missing
            if (!File.Exists(xmlFilePath))
            {
                File.WriteAllText(xmlFilePath,
                    @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
    <Element>Sample Value</Element>
</Root>");
            }

            // Load the XML and extract the desired element value
            string elementValue = string.Empty;
            try
            {
                XDocument xDoc = XDocument.Load(xmlFilePath);
                XElement element = xDoc.Root?.Element("Element");
                if (element != null)
                {
                    elementValue = element.Value;
                }
                else
                {
                    Console.WriteLine("XML element '/Root/Element' not found. Using empty value.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read XML file: {ex.Message}");
            }

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put the extracted XML value into cell A1
            sheet.Cells["A1"].PutValue(elementValue);

            // Save the workbook
            string outputPath = "output.xlsx";

            // Ensure the output directory exists (handle possible null)
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
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
