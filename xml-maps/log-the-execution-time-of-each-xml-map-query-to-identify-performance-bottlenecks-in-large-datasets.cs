// Title: C# – Log XmlMapQuery Execution Time in Aspose.Cells to Spot Performance Bottlenecks
// Description: This example creates a workbook, imports XML, retrieves the first XmlMap, runs XPath queries with Worksheet.XmlMapQuery, measures each query with Stopwatch, prints the path, elapsed milliseconds and cell‑area count, and saves the file. Use it to benchmark XML map queries on large data sets and identify slow paths.
// Keywords: Aspose.Cells XmlMapQuery performance | measure XmlMap query time C# | benchmark XML map queries .NET | stopwatch Aspose.Cells profiling | XML map query latency | large XML dataset performance Aspose | C# performance logging Aspose.Cells
// Common Searches: how to time XmlMapQuery in Aspose.Cells | profile XML map query speed C# | measure execution time of Worksheet.XmlMapQuery | Aspose.Cells performance testing for XML maps | log XmlMap query duration .NET
// Developer Intent: Measure and record the execution time of each XmlMap query to detect slow XPath paths in large XML datasets.
// Use Cases: Quickly benchmark a single XPath query against an XmlMap and view the elapsed time. | Iterate over multiple XPath expressions, log each duration, and compare results to find the most expensive queries. | Integrate timing data into automated tests or CI pipelines to monitor regression in XML map query performance.
// AI Prompts: Generate C# code that runs several XmlMapQuery paths, records execution times, and exports the results to a CSV file. | Explain how Stopwatch can be combined with Worksheet.XmlMapQuery to benchmark large XML imports in Aspose.Cells. | Provide optimization recommendations for slow XmlMap queries based on measured execution times in a .NET application.

using System;
using System.Collections;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// This example creates a workbook, imports XML, retrieves the first XmlMap, runs XPath queries with Worksheet.XmlMapQuery, measures each query with Stopwatch, prints the path, elapsed milliseconds and cell‑area count, and saves the file. Use it to benchmark XML map queries on large data sets and identify slow paths.
class XmlMapQueryPerformanceDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Sample XML data to import
        string xmlData = @"<?xml version='1.0' encoding='UTF-8'?>
            <Root>
                <Data>
                    <Item>Value1</Item>
                    <Item>Value2</Item>
                </Data>
            </Root>";

        // Import XML into the workbook (uses the provided ImportXml method)
        using (MemoryStream ms = new MemoryStream())
        using (StreamWriter writer = new StreamWriter(ms))
        {
            writer.Write(xmlData);
            writer.Flush();
            ms.Position = 0;
            workbook.ImportXml(ms, "Sheet1", 0, 0);
        }

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one XmlMap
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XmlMap found in the workbook.");
            return;
        }

        // Get the first XmlMap
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the XML paths to query
        string[] queryPaths = new string[]
        {
            "/Root/Data/Item"
        };

        // Execute each query and log its execution time
        foreach (string path in queryPaths)
        {
            Stopwatch timer = Stopwatch.StartNew();

            // Perform the query (uses the provided XmlMapQuery method)
            ArrayList cellAreas = worksheet.XmlMapQuery(path, xmlMap);

            timer.Stop();

            Console.WriteLine($"Query Path: {path}");
            Console.WriteLine($"Execution Time: {timer.ElapsedMilliseconds} ms");
            Console.WriteLine($"Number of Cell Areas Returned: {cellAreas.Count}");
        }

        // Save the workbook (uses the standard Save method)
        workbook.Save("XmlMapQueryPerformance.xlsx");
    }
}
