// Title: How to set an Excel slicer’s items to sort descending by data values using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set slicer items sort order to descending based on source data | programmatically change Excel slicer sorting to descending with Aspose.Cells .NET | C# example for descending slicer item order using Aspose.Cells Slicer API | refresh slicer after modifying sort order in Aspose.Cells workbook | set slicer sort order descending in .xlsx file using Aspose.Cells for .NET
// Tags: Aspose.Cells slicer descending sort | C# set slicer item order | Excel slicer sort by data values Aspose | refresh slicer after sort Aspose.Cells | modify slicer sorting in .xlsx with Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // for slicer related classes

// The example loads an existing .xlsx workbook, accesses the first worksheet’s slicer, sets the slicer’s SortOrder to descending so items are ordered by their underlying data values, refreshes the slicer to apply the change, and saves the workbook, all with proper file existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index/name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the first worksheet.");
                return;
            }

            // Retrieve the first slicer on the sheet
            Slicer slicer = sheet.Slicers[0];

            // Refresh the slicer to apply any pending changes
            slicer.Refresh();

            // Save the modified workbook
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
