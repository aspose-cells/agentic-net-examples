// Title: How to remove all slicers from every worksheet in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# remove slicers from all worksheets in a workbook | clear slicer collection for each sheet using Aspose.Cells .NET | delete every slicer in an Excel file programmatically with Aspose.Cells | iterate worksheets and clear slicers Aspose.Cells example | C# code to remove slicers from Excel workbook using Aspose.Cells
// Tags: Aspose.Cells workbook slicer removal | C# clear worksheet slicers Aspose | Iterate worksheets delete slicers Aspose.Cells | Excel slicer collection clear .NET | Workbook.Slicers.Clear method Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, iterates through each worksheet, clears all slicers via the worksheet's Slicers collection, and saves the modified file, demonstrating error handling and the Workbook.Slicers.Clear API in Aspose.Cells for .NET.
class RemoveSlicersExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and remove its slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear all slicers on the current worksheet
                sheet.Slicers.Clear();
            }

            // Save the modified workbook to the output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
