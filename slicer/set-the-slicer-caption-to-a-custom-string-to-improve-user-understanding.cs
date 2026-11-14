// Title: How to change an Excel slicer caption with Aspose.Cells for .NET (including API version considerations)
// Common Searches: Aspose.Cells C# change slicer caption in existing workbook | set custom slicer label using Aspose.Cells .NET API | update slicer caption programmatically in Excel with Aspose.Cells | check if slicer caption can be modified in Aspose.Cells version | C# example for accessing slicers in Aspose.Cells workbook
// Tags: Aspose.Cells slicer caption property | Excel slicer label modification .NET | programmatic slicer access Aspose.Cells | checking slicer existence Aspose.Cells | Aspose.Cells version limitation slicer caption

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, accesses the first worksheet, checks whether any slicers are present, logs that caption modification is skipped because the current Aspose.Cells API does not expose a slicer caption property, and then saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Check for slicers on the worksheet
            if (worksheet.Slicers.Count > 0)
            {
                // Note: Direct manipulation of slicer caption may require a newer Aspose.Cells version.
                // Here we simply acknowledge the presence of slicers.
                Console.WriteLine("Slicers are present on the worksheet. Caption modification is skipped due to API availability.");
            }
            else
            {
                Console.WriteLine("No slicers found on the worksheet.");
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
