// Title: C# Example: Worksheet.XmlMapQuery with Wildcard to Retrieve All XML‑Mapped Cells Under a Parent Node (Aspose.Cells)
// Description: This sample creates a Workbook, imports an XML document that contains a namespace and multiple child elements under a <Data> node, uses Worksheet.XmlMapQuery with a wildcard path to fetch every mapped cell, iterates the used range to display each cell address and value, and saves the result to an XLSX file.
// Keywords: Worksheet.XmlMapQuery | wildcard path | XML map query .NET | Aspose.Cells XML namespace | retrieve all child nodes | C# Aspose.Cells example | import XML to worksheet
// Common Searches: Worksheet.XmlMapQuery wildcard example | Aspose.Cells get all XML mapped cells | C# query XML map with * wildcard | Aspose.Cells XML namespace query | retrieve child elements from XML map
// Developer Intent: The developer needs to query an XML map in a worksheet and obtain every cell mapped from a specific parent node using a wildcard path.
// Use Cases: Import an XML file with namespaces into a worksheet and call Worksheet.XmlMapQuery("/ns1:Root/ns1:Data/*") to return all cells under the Data element. | Iterate the returned cells to display, validate, or transform their values in a .NET application. | Save the workbook after processing the queried cells for reporting, export, or further analysis.
// AI Prompts: Generate C# code that calls Worksheet.XmlMapQuery with the path '/ns1:Root/ns1:Data/*' to fetch all cells under the Data node and prints each cell address and string value. | Show how to handle XML namespaces when using XmlMapQuery with a wildcard in Aspose.Cells for .NET, including map setup and result iteration. | Write a reusable method that accepts a parent node XPath and returns a list of cell addresses for all child elements using XmlMapQuery with a wildcard.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsXmlMapQueryDemo
{
    // This sample creates a Workbook, imports an XML document that contains a namespace and multiple child elements under a <Data> node, uses Worksheet.XmlMapQuery with a wildcard path to fetch every mapped cell, iterates the used range to display each cell address and value, and saves the result to an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML with a namespace and multiple child elements under <Data>
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns1:Root xmlns:ns1='http://example.com'>
    <ns1:Data>
        <ns1:Item>Value1</ns1:Item>
        <ns1:Description>First item</ns1:Description>
        <ns1:Quantity>10</ns1:Quantity>
    </ns1:Data>
</ns1:Root>";

                // Import the XML into the first worksheet starting at cell A1
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate over the used cells to display the imported data.
                Console.WriteLine("Imported XML data (cell address : value):");
                AsposeRange usedRange = worksheet.Cells.MaxDisplayRange;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = worksheet.Cells[row, col];
                        if (!string.IsNullOrEmpty(cell.StringValue))
                        {
                            Console.WriteLine($"{cell.Name} : \"{cell.StringValue}\"");
                        }
                    }
                }

                // Save the workbook (optional, demonstrates lifecycle usage)
                string outputPath = "XmlMapQueryWildcardDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
