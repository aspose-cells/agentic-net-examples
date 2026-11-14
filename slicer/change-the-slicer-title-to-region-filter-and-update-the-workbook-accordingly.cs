// Title: Change slicer caption to "Region Filter" in an existing Excel workbook using Aspose.Cells for .NET
// Common Searches: c# update slicer title to Region Filter in an existing Excel file with Aspose.Cells | programmatically set slicer caption across all worksheets using Aspose.Cells .NET | how to modify slicer caption in a loaded workbook with Aspose.Cells for C# | Aspose.Cells example to change slicer caption in Excel workbook | set custom slicer caption in Excel using Aspose.Cells API
// Tags: Aspose.Cells slicer caption property | C# update slicer title in Excel | programmatic slicer caption change .NET | modify slicer caption across worksheets | set slicer caption for existing workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The sample loads an existing Excel workbook, iterates through each worksheet's slicer collection, sets every slicer's Caption property to "Region Filter", and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and update slicer captions
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;
                if (slicers != null && slicers.Count > 0)
                {
                    foreach (Slicer slicer in slicers)
                    {
                        // Set the slicer title to "Region Filter"
                        slicer.Caption = "Region Filter";
                    }
                }
            }

            // Save the workbook with the updated slicer titles
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
