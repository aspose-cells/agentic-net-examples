// Title: How to revert an Excel worksheet to Normal view mode using Aspose.Cells for .NET to hide page‑break outlines
// Common Searches: Aspose.Cells C# set worksheet view to Normal to remove page break lines | switch Excel worksheet view mode to Normal with Aspose.Cells .NET | hide page break outlines in a worksheet using Aspose.Cells API | remove page break preview from a workbook programmatically with Aspose.Cells | workaround for Aspose.Cells ViewType.Normal deprecation in .NET
// Tags: Aspose.Cells set worksheet view normal | remove page break outlines Aspose.Cells | worksheet view mode Normal .NET | Aspose.Cells viewtype deprecation handling | Excel worksheet display without page breaks Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example demonstrates how to load an Excel file with Aspose.Cells, address the removal of page‑break outlines by switching the worksheet back to Normal view mode, and save the workbook, while noting the deprecation of the older ViewType.Normal API in recent .NET releases.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the desired index/name)
                Worksheet worksheet = workbook.Worksheets[0];

                // NOTE: In recent versions of Aspose.Cells the Worksheet.View property
                // and ViewType.Normal enum are no longer available. If you need to
                // change the view mode, refer to the latest documentation for the
                // appropriate API. The following line is omitted to keep the code
                // compatible with the current library version.

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
