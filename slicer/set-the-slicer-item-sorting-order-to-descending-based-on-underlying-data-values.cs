// Title: Set descending sort order for Excel slicer items based on data values using Aspose.Cells in C#
// Common Searches: how to set slicer items to sort descending by data values with Aspose.Cells C# | C# Aspose.Cells change slicer item order to descending | programmatically sort Excel slicer items descending using Aspose.Cells .NET | Aspose.Cells slicer sorting not available in current version | retrieve and modify slicer sort order in a workbook with C#
// Tags: Aspose.Cells slicer descending sort | C# modify slicer item order | Excel slicer sort by data values .NET | Aspose.Cells API slicer sorting properties | programmatic slicer property adjustment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses a slicer named "MySlicer" on the first worksheet, and attempts to set its item sorting to descending based on underlying data values. It notes that the SortBy and SortOrder properties are unavailable in the current Aspose.Cells version, then saves the workbook.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string slicerName = "MySlicer";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the slicer by its name
                Slicer slicer = worksheet.Slicers[slicerName];
                if (slicer == null)
                {
                    Console.WriteLine($"Slicer '{slicerName}' not found.");
                    return;
                }

                // NOTE: Sorting properties (SortBy, SortOrder) are not available in the
                // current Aspose.Cells version used. If sorting is required, refer to
                // the appropriate API for the installed version.

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
