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

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Sample XML that will be used to create an XML map
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
                <ns1:Root xmlns:ns1='http://example.com'>
                    <ns1:Data>
                        <ns1:Item>Value1</ns1:Item>
                        <ns1:Item>Value2</ns1:Item>
                    </ns1:Data>
                </ns1:Root>";

                // Import the XML into the worksheet to create the map
                // This will also place the first mapped value into the sheet
                workbook.ImportXml(xml, "Sheet1", 0, 0);

                // Retrieve the created XML map
                XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

                // Optionally link another cell to the same XML path for demonstration
                worksheet.Cells.LinkToXmlMap(xmlMap.Name, 5, 0, "/ns1:Root/ns1:Data/ns1:Item");

                // Define the XML path we want to query
                string queryPath = "/ns1:Root/ns1:Data/ns1:Item";

                // Query the worksheet for all cell areas linked to the specified XML path
                ArrayList cellAreas = worksheet.XmlMapQuery(queryPath, xmlMap);

                // Output the results
                Console.WriteLine($"Number of cell areas linked to path '{queryPath}': {cellAreas.Count}");
                foreach (CellArea area in cellAreas)
                {
                    // For each area, display its start row/column (zero‑based) and the cell value
                    int row = area.StartRow;
                    int col = area.StartColumn;
                    string cellAddress = CellsHelper.CellIndexToName(row, col);
                    string cellValue = worksheet.Cells[row, col].StringValue;
                    Console.WriteLine($"Cell {cellAddress}: Value = '{cellValue}'");
                }

                // Save the workbook (optional, just to illustrate lifecycle usage)
                workbook.Save("XmlMapQueryDemo.xlsx");
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