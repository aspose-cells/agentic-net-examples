// Title: Transform XML node values (trim whitespace, uppercase) before mapping them to Excel cells with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XDocument, trims whitespace and converts each <Name> element to uppercase, then writes the values into an Aspose.Cells worksheet. | Create a reusable C# method that receives an XDocument, applies a custom transformation to selected XML nodes, and populates an Excel workbook using Aspose.Cells. | Show an example of mapping transformed XML rows to cells in a new workbook and saving the file with Aspose.Cells for .NET.
// Common Searches: asp.net apply custom transformation to XML nodes before exporting to Excel with Aspose.Cells | c# trim whitespace and uppercase XML element values when using Aspose.Cells XML mapping | example loading XDocument, modifying node values, and writing to worksheet using Aspose.Cells
// Tags: xml node transformation with Aspose.Cells | custom value mapping from XDocument to Excel worksheet | trim whitespace and uppercase XML elements in C# | populate Excel cells from transformed XML using Aspose.Cells | save transformed XML data to .xlsx with Aspose.Cells

using System;
using System.Xml.Linq;
using Aspose.Cells;
using System.IO;

// The example loads XML data into an XDocument, trims whitespace and converts each <Name> element to uppercase, then iterates through the transformed nodes and writes the values row‑by‑row into an Aspose.Cells worksheet, finally saving the workbook as TransformedData.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Example XML data that will be mapped to cells
            string xmlData = @"<Root>
                <Row>
                    <Name>John Doe</Name>
                    <Age>30</Age>
                </Row>
                <Row>
                    <Name>Jane Smith</Name>
                    <Age>25</Age>
                </Row>
            </Root>";

            // Load the XML into an XDocument for manipulation
            XDocument doc = XDocument.Parse(xmlData);

            // Custom transformation: trim whitespace and convert names to upper case
            foreach (XElement nameElement in doc.Descendants("Name"))
            {
                string original = nameElement.Value;
                string transformed = original.Trim().ToUpperInvariant();
                nameElement.Value = transformed;
            }

            // Write transformed XML data into worksheet cells
            Worksheet sheet = workbook.Worksheets[0];
            int rowIndex = 0;
            foreach (XElement row in doc.Descendants("Row"))
            {
                int colIndex = 0;
                foreach (XElement cell in row.Elements())
                {
                    sheet.Cells[rowIndex, colIndex].PutValue(cell.Value);
                    colIndex++;
                }
                rowIndex++;
            }

            // Define output file path
            string outputPath = "TransformedData.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
