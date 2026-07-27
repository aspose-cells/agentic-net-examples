using System;
using System.Collections;
using System.Diagnostics;
using Aspose.Cells;

namespace XmlMapQueryPerformanceDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Sample XML data to import and create a map
            string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
                <ns1:Root xmlns:ns1='http://example.com'>
                    <ns1:Data>
                        <ns1:Item>Value1</ns1:Item>
                        <ns1:Item>Value2</ns1:Item>
                        <ns1:Item>Value3</ns1:Item>
                    </ns1:Data>
                </ns1:Root>";

            // Import the XML into the first worksheet starting at cell A1
            workbook.ImportXml(xmlData, "Sheet1", 0, 0);

            // Retrieve the first (and only) XmlMap from the workbook
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XmlMap found after import.");
                return;
            }

            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Access the worksheet where the map resides
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the XML paths we want to query
            string[] paths = new string[]
            {
                "/ns1:Root/ns1:Data/ns1:Item",
                "/ns1:Root/ns1:Data",
                "/ns1:Root"
            };

            // Iterate over each path, measure execution time, and log results
            foreach (string path in paths)
            {
                Stopwatch sw = Stopwatch.StartNew();

                // Perform the XmlMapQuery
                ArrayList cellAreas = worksheet.XmlMapQuery(path, xmlMap);

                sw.Stop();

                // Log execution time
                Console.WriteLine($"Query Path: {path}");
                Console.WriteLine($"Execution Time: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine($"Returned Areas: {cellAreas.Count}");

                // Optionally display the first cell value for verification
                if (cellAreas.Count > 0)
                {
                    CellArea area = (CellArea)cellAreas[0];
                    string cellValue = worksheet.Cells[area.StartRow, area.StartColumn].StringValue;
                    Console.WriteLine($"First Mapped Cell: Row {area.StartRow}, Column {area.StartColumn}, Value \"{cellValue}\"");
                }

                Console.WriteLine(new string('-', 50));
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("XmlMapQueryPerformanceDemo.xlsx");
        }
    }
}