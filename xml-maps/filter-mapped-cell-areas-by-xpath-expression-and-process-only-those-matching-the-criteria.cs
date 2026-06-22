using System;
using System.Collections;
using Aspose.Cells;

class FilterMappedCellAreas
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Sample XML containing a namespace and two <Item> elements
        string xml = @"<?xml version='1.0' encoding='UTF-8'?>
            <ns1:Root xmlns:ns1='http://example.com'>
                <ns1:Data>
                    <ns1:Item>Value1</ns1:Item>
                    <ns1:Item>Value2</ns1:Item>
                </ns1:Data>
            </ns1:Root>";

        // Import the XML into the worksheet; this creates an XML map automatically
        workbook.ImportXml(xml, worksheet.Name, 0, 0);

        // Retrieve the first (and only) XML map created by ImportXml
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // XPath expression that points to the XML elements we want to process
        string xpath = "/ns1:Root/ns1:Data/ns1:Item";

        // Query the worksheet for cell areas that are mapped to the specified XPath
        ArrayList mappedAreas = worksheet.XmlMapQuery(xpath, xmlMap);

        // Apply an auto‑filter to each cell area that matches the XPath
        foreach (CellArea area in mappedAreas)
        {
            worksheet.Filter(area);
        }

        // Save the workbook with the applied filters
        workbook.Save("FilteredMappedAreas.xlsx");
    }
}