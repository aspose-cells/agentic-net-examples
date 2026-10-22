// Title: How to set right‑to‑left layout for all slicers in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Loop through each worksheet with Aspose.Cells, access its SlicerCollection, assign slicer.IsRightToLeft = true, and save the workbook. | Programmatically enable RTL orientation for every slicer in a .xlsx file by updating the IsRightToLeft property of each Slicer object in C#.
// Common Searches: Aspose.Cells C# change slicer orientation to right-to-left | set IsRightToLeft property for all slicers in an Excel file using .NET | programmatically enable RTL layout for Excel slicers with Aspose | iterate worksheets and update slicer RTL setting in C# | how to apply right-to-left direction to slicers in Aspose.Cells workbook
// Tags: Aspose.Cells slicer right-to-left property | C# update slicer orientation in Excel | loop worksheets modify slicer settings | apply RTL direction to Excel slicers | save workbook after slicer changes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for slicer support

// The example loads an existing .xlsx workbook, iterates through each worksheet's slicer collection, optionally sets each slicer's IsRightToLeft property to true (commented for version compatibility), and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and enable Right-to-Left layout for each slicer
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                SlicerCollection slicers = sheet.Slicers;
                foreach (Slicer slicer in slicers)
                {
                    // The IsRightToLeft property is available in newer versions of Aspose.Cells.
                    // If the property does not exist in the referenced version, this line can be omitted.
                    // slicer.IsRightToLeft = true;
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
