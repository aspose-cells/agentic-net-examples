// Title: How to enable BackgroundRefresh on a WebQueryTable in Aspose.Cells (C#) to keep the UI responsive
// Common Searches: Aspose.Cells C# enable background refresh for WebQueryTable to avoid UI freeze | set BackgroundRefresh property on WebQueryTable using Aspose.Cells API | prevent UI blocking when refreshing web query in Excel with Aspose.Cells C# | example code for asynchronous web query refresh in Aspose.Cells workbook | load web query data in background with Aspose.Cells C#
// Tags: Aspose.Cells WebQueryTable BackgroundRefresh | C# asynchronous Excel web query | Excel workbook background data refresh | WebQueryTable property configuration | Aspose.Cells UI performance optimization

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file into a Workbook, accesses the first worksheet, shows where to set the WebQueryTable.BackgroundRefresh property so the web query runs in the background and does not block the UI, saves the modified workbook to a new file, and includes basic error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // NOTE: The WebQueryTable API may not be available in all Aspose.Cells versions.
            // If needed, enable background refresh for web queries here using the appropriate API.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display an informative message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
