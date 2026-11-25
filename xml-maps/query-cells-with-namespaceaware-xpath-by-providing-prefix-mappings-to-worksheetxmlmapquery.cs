// Title: Retrieve cell ranges mapped to a namespace‑aware XPath using Worksheet.GetAreasOfXmlMapQuery in Aspose.Cells for .NET
// Common Searches: how to query cells with a namespace‑aware XPath using Aspose.Cells C# | using GetAreasOfXmlMapQuery to find mapped cells for an XML namespace prefix | retrieve cell addresses for /ns:Root/ns:Data/ns:Item after ImportXml in .NET | example of namespace‑aware XML map query in Aspose.Cells workbook | C# code to get cell areas linked to a namespaced XPath in Aspose.Cells
// Tags: Worksheet.GetAreasOfXmlMapQuery XML namespace support | Aspose.Cells XML map cell area extraction | XML import with namespaces Aspose.Cells | C# XML namespace XPath mapping Aspose.Cells | mapped cell ranges for XPath Aspose.Cells

using System;
using Aspose.Cells;
using System.Collections.Generic;

namespace XmlMapNamespaceQueryDemo
{
    // The example imports an XML string that uses a namespace prefix into a new workbook, automatically creates an XmlMap, defines a matching namespace‑aware XPath, and calls Worksheet.GetAreasOfXmlMapQuery to obtain the cell areas linked to that XPath. It iterates through each cell in the returned areas, prints the address and value, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Prepare a sample XML that uses a namespace prefix (ns)
            string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns:Root xmlns:ns='http://example.com'>
    <ns:Data>
        <ns:Item>Value1</ns:Item>
        <ns:Item>Value2</ns:Item>
        <ns:Item>Value3</ns:Item>
    </ns:Data>
</ns:Root>";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Import the XML into the first worksheet.
            // This automatically creates an XmlMap for the imported XML.
            // The data will be placed starting at cell A1.
            workbook.ImportXml(xml, "Sheet1", 0, 0);

            // Get references to the first worksheet and the generated XmlMap
            Worksheet worksheet = workbook.Worksheets[0];
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XmlMap was created during ImportXml.");
                return;
            }
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the namespace‑aware XPath expression.
            // The prefix "ns" is the same as declared in the XML (http://example.com).
            string xpath = "/ns:Root/ns:Data/ns:Item";

            // Query the worksheet for cell areas that are linked to the specified XPath.
            // GetAreasOfXmlMapQuery is the recommended (non‑obsolete) API.
            CellArea[] areas = worksheet.GetAreasOfXmlMapQuery(xpath, xmlMap);

            // Output the results
            if (areas.Length == 0)
            {
                Console.WriteLine($"No cells were mapped to the XPath '{xpath}'.");
            }
            else
            {
                Console.WriteLine($"Found {areas.Length} cell area(s) for XPath '{xpath}':");
                foreach (CellArea area in areas)
                {
                    // The query may return multi‑cell ranges; iterate through each cell in the area.
                    for (int row = area.StartRow; row <= area.EndRow; row++)
                    {
                        for (int col = area.StartColumn; col <= area.EndColumn; col++)
                        {
                            Cell cell = worksheet.Cells[row, col];
                            string address = CellsHelper.CellIndexToName(row, col);
                            Console.WriteLine($"{address}: {cell.StringValue}");
                        }
                    }
                }
            }

            // Save the workbook so you can inspect the mapping in Excel if desired.
            string outFile = "XmlMapNamespaceQueryResult.xlsx";
            workbook.Save(outFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outFile}'.");
        }
    }
}
