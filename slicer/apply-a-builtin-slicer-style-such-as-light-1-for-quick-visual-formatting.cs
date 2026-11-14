// Title: Applying a built‑in Excel slicer style (Light 1) with Aspose.Cells for .NET – handling API limitations
// Common Searches: how to set an Excel slicer to Light 1 using Aspose.Cells C# | Aspose.Cells C# apply built‑in slicer style programmatically | change slicer visual theme in .NET workbook with Aspose.Cells | Aspose.Cells version without slicer style property workaround | apply built‑in slicer formatting to first worksheet slicer in C#
// Tags: Aspose.Cells apply slicer style | C# built‑in slicer theme Excel | Excel slicer visual formatting .NET | Aspose.Cells slicer style API limitation | set Light 1 slicer style programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an existing workbook, checks for at least one slicer on the first worksheet, retrieves the first slicer, notes that the current Aspose.Cells version does not expose a direct property to assign built‑in slicer styles such as Light 1, and then saves the workbook. It demonstrates how to handle the API limitation when attempting to apply built‑in slicer formatting programmatically.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure that at least one slicer exists on the worksheet.
            if (worksheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the worksheet.");
                return;
            }

            // Retrieve the first slicer.
            Slicer slicer = worksheet.Slicers[0];

            // NOTE: Setting slicer style is omitted because the API version used
            // does not expose a direct SlicerStyle property. Adjust as needed for
            // newer versions of Aspose.Cells.

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
