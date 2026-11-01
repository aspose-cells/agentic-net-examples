// Title: Hide slicer header row in an Excel worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide slicer header row in Excel | remove slicer title from worksheet using Aspose.Cells .NET | set slicer IsHeaderVisible property false Aspose.Cells | how to hide slicer caption in generated Excel file C# | compact slicer control without header using Aspose.Cells | Aspose.Cells hide slicer header when property not exposed
// Tags: slicer header visibility Aspose.Cells | remove slicer caption .NET | Excel slicer formatting with Aspose.Cells | reflection modify slicer properties C# | compact slicer control without title Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, checks for slicers on the first worksheet, and uses reflection to set the slicer's IsHeaderVisible property to false (when available) to hide the header row, then saves the workbook.
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

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // If a slicer exists, attempt to hide its header (if supported)
            if (sheet.Slicers.Count > 0)
            {
                var slicer = sheet.Slicers[0];
                var prop = slicer.GetType().GetProperty("IsHeaderVisible");
                if (prop != null && prop.CanWrite)
                {
                    // Hide the slicer header
                    prop.SetValue(slicer, false);
                }
                else
                {
                    // Header visibility not supported; optionally log or ignore
                    Console.WriteLine("IsHeaderVisible property not available for this slicer.");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
