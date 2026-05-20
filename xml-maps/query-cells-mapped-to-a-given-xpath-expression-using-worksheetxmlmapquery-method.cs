using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapQueryDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Sample XML data to import (creates an XML map automatically)
                string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
<ns1:Root xmlns:ns1='http://example.com'>
    <ns1:Data>
        <ns1:Item>Value1</ns1:Item>
        <ns1:Item>Value2</ns1:Item>
    </ns1:Data>
</ns1:Root>";

                // Import the XML into the first worksheet starting at cell A1
                workbook.ImportXml(xmlData, "Sheet1", 0, 0);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the XML map that was created during ImportXml
                if (workbook.Worksheets.XmlMaps.Count == 0)
                {
                    Console.WriteLine("No XML maps were created.");
                    return;
                }
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // (Optional) Link a specific cell to the XML path.
                worksheet.Cells.LinkToXmlMap(xmlMap.Name, 0, 0, "/ns1:Root/ns1:Data/ns1:Item");

                // Query the worksheet for cell areas mapped to the given XPath
                string xpath = "/ns1:Root/ns1:Data/ns1:Item";
                ArrayList cellAreas = worksheet.XmlMapQuery(xpath, xmlMap);

                // Output the results
                if (cellAreas.Count > 0)
                {
                    foreach (CellArea area in cellAreas)
                    {
                        Console.WriteLine($"Mapped cell at Row {area.StartRow}, Column {area.StartColumn}");
                        Console.WriteLine($"Cell Value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
                    }
                }
                else
                {
                    Console.WriteLine("No cells are mapped to the specified XPath.");
                }

                // Save the workbook (lifecycle rule: save)
                string outputPath = "XmlMapQueryResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            XmlMapQueryDemo.Run();
        }
    }
}