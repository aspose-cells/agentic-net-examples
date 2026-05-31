using System;
using System.Collections;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsNamespaceAwareXPathDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample XML with a namespace declaration
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<ns1:Root xmlns:ns1='http://example.com'>
    <ns1:Data>
        <ns1:Item>Value1</ns1:Item>
        <ns1:Item>Value2</ns1:Item>
    </ns1:Data>
</ns1:Root>";

                // Import the XML into the worksheet and create an XML map.
                // Use the overload that accepts a Stream (no bool parameter needed).
                using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
                {
                    workbook.ImportXml(ms, "Sheet1", 0, 0);
                }

                // Ensure that an XML map was created
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML map was created.");
                    return;
                }

                // Retrieve the first (and only) XML map from the workbook
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Namespace‑aware XPath query using the prefix defined in the XML (ns1)
                string xpath = "/ns1:Root/ns1:Data/ns1:Item";

                // Query the worksheet for cell areas linked to the specified XML path
                ArrayList cellAreas = worksheet.XmlMapQuery(xpath, xmlMap);

                // Output the results
                if (cellAreas.Count > 0)
                {
                    foreach (CellArea area in cellAreas)
                    {
                        Console.WriteLine($"Found data at Row {area.StartRow + 1}, Column {area.StartColumn + 1}");
                        Console.WriteLine($"Cell value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
                    }
                }
                else
                {
                    Console.WriteLine("No cells were mapped to the specified XPath.");
                }

                // Save the workbook (optional, just to demonstrate lifecycle usage)
                string outputPath = "NamespaceAwareXPathResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}