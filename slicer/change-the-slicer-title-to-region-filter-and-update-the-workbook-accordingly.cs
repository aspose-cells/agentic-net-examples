// Title: Change slicer caption to 'Region Filter' in an existing Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# snippet that opens an Excel file with Aspose.Cells, iterates over all slicers on the first worksheet, sets each slicer's Caption property to "Region Filter", and saves the workbook. | Generate code that loads a workbook, updates the title of every slicer to a custom text, handles a missing input file gracefully, and writes the result to a new file using Aspose.Cells.
// Common Searches: asp.net change slicer caption Aspose.Cells | C# update slicer title in existing Excel file using Aspose.Cells | how to set slicer caption programmatically with Aspose.Cells .NET | rename all slicers on first worksheet to custom text Aspose.Cells | Aspose.Cells slicer caption property example C#
// Tags: Aspose.Cells modify slicer title | C# change Excel slicer caption | Aspose.Cells set slicer caption .NET | programmatic Excel slicer rename C# | Excel workbook slicer title update Aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer class

// The program loads 'input.xlsx', iterates through every slicer on the first worksheet, sets each slicer's Caption to "Region Filter", and saves the modified workbook as 'output.xlsx', with error handling for missing files.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            var worksheet = workbook.Worksheets[0];

            // Update the caption of each slicer on the worksheet
            foreach (Slicer slicer in worksheet.Slicers)
            {
                slicer.Caption = "Region Filter";
            }

            // Save the modified workbook
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
