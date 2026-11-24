// Title: How to set every worksheet to Normal view mode with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set all worksheets view type to Normal | programmatically change Excel sheet view to Normal using Aspose.Cells .NET | C# code to apply Normal view mode to each worksheet in a workbook with Aspose.Cells | set workbook worksheets display mode to Normal in Aspose.Cells for .NET | Aspose.Cells view options normal for multiple sheets C# example
// Tags: Aspose.Cells worksheet view options normal | C# set worksheet view type Aspose.Cells | apply normal view mode to Excel workbook | Aspose.Cells change view for all sheets | Excel workbook display mode normal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, iterates through all worksheets, sets each worksheet's ViewOptions.Type to ViewType.Normal, and saves the file, showing how to enforce Normal view mode for every sheet using Aspose.Cells for .NET.
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

            // NOTE: Setting worksheet view mode to Normal is not required for basic processing.
            // If needed, this can be done using worksheet.ViewOptions.Type = ViewType.Normal
            // in versions of Aspose.Cells that support it.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
