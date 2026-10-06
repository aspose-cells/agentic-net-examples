// Title: Generate sequentially versioned Excel workbooks after modifying each slicer caption with Aspose.Cells for .NET
// AI Prompts: Iterate over all slicers in a workbook, change each slicer's Caption property, and save the workbook to a uniquely numbered file (e.g., Output_v1.xlsx, Output_v2.xlsx) using Aspose.Cells in C#. | Create a C# routine that applies a modification to a slicer, then persists the workbook with an incremental filename pattern to track changes with Aspose.Cells.
// Common Searches: Aspose.Cells C# save workbook after each slicer update with versioned filenames | how to create incremental Excel file names when changing slicer captions in .NET | loop through slicers and export separate workbook versions using Aspose.Cells | C# generate Output_v1.xlsx Output_v2.xlsx for each slicer modification Aspose
// Tags: slicer caption modification Aspose.Cells | incremental workbook versioning C# | save workbook after slicer change Aspose | generate sequential Excel filenames .NET | track slicer edits with versioned files

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The example loads an Excel workbook, loops through each slicer on the first worksheet, updates the slicer's caption, and saves the workbook after each change to a uniquely numbered file (Output_v1.xlsx, Output_v2.xlsx, etc.), demonstrating incremental versioning of workbooks using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume slicers are on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            SlicerCollection slicers = sheet.Slicers;

            // Counter for versioned file names
            int version = 1;

            // Iterate through each slicer, apply a modification, and save the workbook
            for (int i = 0; i < slicers.Count; i++)
            {
                // Access the current slicer
                Slicer slicer = slicers[i];

                // Example modification: change the slicer's caption
                slicer.Caption = $"Modified Slicer {i + 1}";

                // Save the workbook with an incremental versioned file name
                string outputPath = $"Output_v{version}.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Saved modified workbook to: {outputPath}");

                version++;
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
