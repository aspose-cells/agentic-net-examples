// Title: Log enumeration start/end timestamps to benchmark rows, cells, and range iteration in Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over Worksheet.Cells.Rows, Worksheet.Cells, and a specific Range while printing ISO‑8601 start and end timestamps and the elapsed milliseconds for each enumeration using Aspose.Cells. | Create a reusable method that accepts any IEnumerator from Aspose.Cells, records DateTime.Now before and after the loop, and returns the elapsed time in milliseconds. | Demonstrate how to save the workbook after the enumeration passes and display the logged timing information in the console.
// Common Searches: how to measure enumeration time of rows in Aspose.Cells C# | log start and end timestamps for cell range iteration using Aspose.Cells .NET | benchmark Worksheet.Cells.GetEnumerator performance Aspose.Cells | record enumeration duration for Aspose.Cells range A1:B100 in C# | track performance of Aspose.Cells enumeration across multiple passes
// Tags: Aspose.Cells enumeration timing | row iteration performance measurement | cell collection duration logging | range enumeration benchmark .NET | C# Aspose.Cells performance logging

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, fills 100 rows with data, and uses a helper method to log ISO‑8601 start and end timestamps plus elapsed milliseconds for enumerating rows, all cells, and the A1:B100 range, then saves the workbook.
class EnumerationPerformanceLogger
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (100 rows, 2 columns)
            for (int i = 0; i < 100; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
                worksheet.Cells[i, 1].PutValue(i);
            }

            // First enumeration: iterate over all rows
            LogEnumeration("Rows", worksheet.Cells.Rows.GetEnumerator());

            // Second enumeration: iterate over all cells in the worksheet
            LogEnumeration("Cells", worksheet.Cells.GetEnumerator());

            // Third enumeration: iterate over a specific range (A1:B100)
            AsposeRange range = worksheet.Cells.CreateRange("A1:B100");
            LogEnumeration("Range", range.GetEnumerator());

            // Save the workbook (standard save operation)
            string outputPath = "EnumerationPerformance.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method that logs start/end timestamps and duration of an enumeration pass
    static void LogEnumeration(string collectionName, IEnumerator enumerator)
    {
        try
        {
            DateTime startTime = DateTime.Now;
            Console.WriteLine($"{collectionName} enumeration started at {startTime:O}");

            while (enumerator.MoveNext())
            {
                // Access the current element to ensure the enumerator advances
                var current = enumerator.Current;
            }

            DateTime endTime = DateTime.Now;
            Console.WriteLine($"{collectionName} enumeration ended at {endTime:O}");

            double durationMs = (endTime - startTime).TotalMilliseconds;
            Console.WriteLine($"{collectionName} duration: {durationMs} ms");
            Console.WriteLine(); // Blank line for readability
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Enumeration error for {collectionName}: {ex.Message}");
        }
    }
}
