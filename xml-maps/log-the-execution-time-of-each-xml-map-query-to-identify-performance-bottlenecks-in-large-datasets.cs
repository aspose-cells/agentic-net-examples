// Title: How to log execution time of each XML map query in Aspose.Cells for .NET using Stopwatch
// Common Searches: Aspose.Cells C# measure GetAreasOfXmlMapQuery execution time | benchmark XML map query performance on large dataset with Aspose.Cells | log execution time of worksheet XML map queries in .NET | profile XML map query speed using Stopwatch in Aspose.Cells | determine bottlenecks of XML map queries in large Excel workbooks
// Tags: GetAreasOfXmlMapQuery execution timing | Aspose.Cells XML map performance profiling | stopwatch based XML map query measurement | large dataset XML map benchmarking Aspose.Cells | worksheet XML map query latency logging

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Aspose.Cells;

// The sample creates a Workbook, generates a large XML document, imports it to automatically create an XmlMap, and then runs several XPath queries using GetAreasOfXmlMapQuery. Each query is timed with a Stopwatch, and the elapsed milliseconds are printed, allowing developers to identify performance bottlenecks in XML map queries on large datasets.
class Program
{
    static void Main()
    {
        // Initialize a new workbook
        Workbook workbook = new Workbook();

        // Create a large XML string with many repeated items to simulate a large dataset
        const int itemCount = 5000; // Adjust this number to increase dataset size
        string xmlData = GenerateSampleXml(itemCount);

        // Import the XML data into the first worksheet.
        // This also creates an XmlMap automatically.
        workbook.ImportXml(xmlData, "Sheet1", 0, 0);

        // Retrieve the worksheet and the XmlMap that was created
        Worksheet worksheet = workbook.Worksheets[0];
        if (workbook.Worksheets.XmlMaps.Count == 0)
        {
            Console.WriteLine("No XmlMap was created during ImportXml.");
            return;
        }
        XmlMap xmlMap = workbook.Worksheets.XmlMaps[0];

        // Define the list of XML element paths we want to query.
        // These paths correspond to the schema used in the generated XML.
        List<string> queryPaths = new List<string>
        {
            "/Root/Data/Item/Name",
            "/Root/Data/Item/Value",
            "/Root/Data/Item"
        };

        Console.WriteLine($"Starting XML map queries on a dataset with {itemCount} items...");
        Console.WriteLine();

        foreach (string path in queryPaths)
        {
            Stopwatch sw = Stopwatch.StartNew();

            // Perform the query using the recommended API (GetAreasOfXmlMapQuery)
            CellArea[] areas = worksheet.GetAreasOfXmlMapQuery(path, xmlMap);

            sw.Stop();

            // Output the results
            Console.WriteLine($"Query Path: {path}");
            Console.WriteLine($"  Areas Returned : {areas?.Length ?? 0}");
            Console.WriteLine($"  Execution Time : {sw.ElapsedMilliseconds} ms");
            Console.WriteLine();
        }

        // Optionally, save the workbook for inspection
        try
        {
            workbook.Save("XmlMapQueryPerformanceDemo.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved as 'XmlMapQueryPerformanceDemo.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }

    /// <param name="itemCount">Number of Item elements to generate.</param>
    /// <returns>XML string.</returns>
    private static string GenerateSampleXml(int itemCount)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine(@"<?xml version='1.0' encoding='UTF-8'?>");
        sb.AppendLine("<Root xmlns='http://example.com'>");
        sb.AppendLine("  <Data>");

        for (int i = 1; i <= itemCount; i++)
        {
            sb.AppendLine("    <Item>");
            sb.AppendLine($"      <Name>Item{i}</Name>");
            sb.AppendLine($"      <Value>{i * 10}</Value>");
            sb.AppendLine("    </Item>");
        }

        sb.AppendLine("  </Data>");
        sb.AppendLine("</Root>");
        return sb.ToString();
    }
}
