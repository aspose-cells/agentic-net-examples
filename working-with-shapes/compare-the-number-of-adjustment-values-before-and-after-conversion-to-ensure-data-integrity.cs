// Title: Check that a chart’s series count stays the same when converting an XLSX workbook to Excel 97‑2003 (.xls) using Aspose.Cells for .NET
// AI Prompts: Load an XLSX workbook, read the NSeries count of its first chart, save the workbook as .xls with Aspose.Cells, reload it, and compare the counts. | Create a C# console program that outputs a warning if the number of chart series changes after saving to Excel 97‑2003 format. | Write code that validates chart data integrity by ensuring the NSeries collection size is identical before and after workbook format conversion.
// Common Searches: how to verify chart series count after converting xlsx to xls with Aspose.Cells | Aspose.Cells ensure chart data integrity when saving workbook as Excel 97-2003 | compare NSeries count before and after workbook format conversion C# | detect chart series loss during XLSX to XLS conversion using Aspose.Cells | C# code to check chart series consistency after saving as .xls
// Tags: chart series count validation Aspose.Cells | XLSX to Excel97-2003 conversion integrity | chart data consistency after format conversion | Aspose.Cells workbook conversion verification | NSeries collection integrity check

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an XLSX workbook, records the first chart's NSeries count, saves the workbook as Excel 97‑2003 (.xls), reloads it, and compares the NSeries count to confirm that the series count remains unchanged after conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string originalPath = "original.xlsx";
            const string convertedPath = "converted.xls";

            // Verify the original workbook exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"File not found: {originalPath}");
                return;
            }

            // Load the original workbook
            Workbook wbOriginal = new Workbook(originalPath);

            // Access the first worksheet and its first chart
            Worksheet wsOriginal = wbOriginal.Worksheets[0];
            if (wsOriginal.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the original worksheet.");
                return;
            }
            Chart chartOriginal = wsOriginal.Charts[0];

            // Use NSeries count as a proxy for chart data integrity
            int beforeCount = chartOriginal.NSeries.Count;

            // Save the workbook to Excel 97-2003 format to simulate conversion
            wbOriginal.Save(convertedPath, SaveFormat.Excel97To2003);

            // Verify the converted workbook was created
            if (!File.Exists(convertedPath))
            {
                Console.WriteLine($"Failed to create converted file: {convertedPath}");
                return;
            }

            // Load the converted workbook
            Workbook wbConverted = new Workbook(convertedPath);
            Worksheet wsConverted = wbConverted.Worksheets[0];
            if (wsConverted.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the converted worksheet.");
                return;
            }
            Chart chartConverted = wsConverted.Charts[0];

            // Compare NSeries counts to ensure data integrity after conversion
            int afterCount = chartConverted.NSeries.Count;

            if (beforeCount == afterCount)
            {
                Console.WriteLine($"Series count matches: {beforeCount}");
            }
            else
            {
                Console.WriteLine($"Series count mismatch. Before: {beforeCount}, After: {afterCount}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
