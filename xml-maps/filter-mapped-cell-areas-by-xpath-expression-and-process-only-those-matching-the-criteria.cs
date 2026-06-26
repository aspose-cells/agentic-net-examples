using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FilterMappedCellAreasDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Sample XML that will be imported as an XML map
            string xml = @"<?xml version='1.0' encoding='UTF-8'?>
                <ns1:Root xmlns:ns1='http://example.com'>
                    <ns1:Data>
                        <ns1:Item>Value1</ns1:Item>
                        <ns1:Item>Value2</ns1:Item>
                    </ns1:Data>
                </ns1:Root>";

            // Import the XML into the worksheet starting at cell A1.
            // This creates an XML map automatically.
            workbook.ImportXml(xml, "Sheet1", 0, 0);

            // Retrieve the created XML map (the first one)
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the XPath expression to locate the desired XML nodes
            string xpath = "/ns1:Root/ns1:Data/ns1:Item";

            // Query the worksheet for cell areas that are mapped to the XPath
            ArrayList mappedAreas = worksheet.XmlMapQuery(xpath, xmlMap);

            // Apply an AutoFilter to each mapped area, if any
            if (mappedAreas.Count > 0)
            {
                foreach (CellArea area in mappedAreas)
                {
                    worksheet.Filter(area);
                }

                Console.WriteLine($"Applied filter to {mappedAreas.Count} mapped cell area(s).");
            }
            else
            {
                Console.WriteLine("No cell areas were mapped to the specified XPath.");
            }

            // Save the workbook to verify the filter was set
            workbook.Save("FilteredMappedAreas.xlsx");
        }
    }
}