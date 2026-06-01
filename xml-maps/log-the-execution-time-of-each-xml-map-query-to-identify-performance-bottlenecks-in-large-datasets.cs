using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class XmlMapQueryPerformanceLogger
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

            // Sample XML data to import
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

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one XmlMap
            if (workbook.Worksheets.XmlMaps.Count == 0)
            {
                Console.WriteLine("No XmlMap found in the workbook.");
                return;
            }
            XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

            // Define the XML paths to query
            string[] paths = new string[]
            {
                "/ns1:Root/ns1:Data/ns1:Item",
                "/ns1:Root/ns1:Data",
                "/ns1:Root"
            };

            // Loop through each path, measure execution time, and log results
            foreach (string path in paths)
            {
                Stopwatch sw = Stopwatch.StartNew();

                // Query cell areas linked to the specific path
                ArrayList cellAreas = worksheet.XmlMapQuery(path, xmlMap);

                sw.Stop();

                Console.WriteLine($"Query Path: {path}");
                Console.WriteLine($"Execution Time: {sw.ElapsedMilliseconds} ms");
                Console.WriteLine($"Returned Areas: {cellAreas.Count}");
                if (cellAreas.Count > 0)
                {
                    CellArea area = (CellArea)cellAreas[0];
                    Console.WriteLine($"First Area - StartRow: {area.StartRow}, StartColumn: {area.StartColumn}");
                    Console.WriteLine($"Cell Value: {worksheet.Cells[area.StartRow, area.StartColumn].StringValue}");
                }
                Console.WriteLine(new string('-', 40));
            }

            // Save the workbook
            string outputPath = "XmlMapQueryPerformanceLog.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}