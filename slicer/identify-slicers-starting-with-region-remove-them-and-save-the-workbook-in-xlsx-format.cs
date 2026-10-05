// Title: Delete slicers whose names start with "Region" from every worksheet and save the workbook as XLSX using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that loads an Excel file, finds all slicers whose Name begins with "Region" on each worksheet, removes them, and writes the result to a new XLSX file. | Show how to iterate through Workbook.Worksheets, collect slicers matching a prefix, delete those slicers, and save the modified workbook using the Aspose.Cells SaveFormat.Xlsx option.
// Common Searches: aspocells c# delete slicers with name starting region | how to remove specific slicers from an Excel workbook using Aspose.Cells | iterate worksheets to find and delete slicers by prefix Aspose.Cells | save workbook after slicer removal Aspose.Cells .NET | example code for removing region slicers with Aspose.Cells
// Tags: remove slicers by name prefix Aspose.Cells | delete region slicers C# Aspose.Cells | iterate worksheets slicer removal Aspose.Cells | save workbook as XLSX Aspose.Cells | Aspose.Cells slicer management example

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.Collections.Generic;
using System.IO;

// Loads 'input.xlsx', removes every slicer whose Name starts with "Region" from all worksheets, and saves the updated file as 'output.xlsx' in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Collect slicers whose names start with "Region"
                List<Slicer> slicersToRemove = new List<Slicer>();
                foreach (Slicer slicer in sheet.Slicers)
                {
                    if (!string.IsNullOrEmpty(slicer.Name) &&
                        slicer.Name.StartsWith("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        slicersToRemove.Add(slicer);
                    }
                }

                // Remove the identified slicers from the worksheet
                foreach (Slicer slicer in slicersToRemove)
                {
                    sheet.Slicers.Remove(slicer);
                }
            }

            // Save the modified workbook in XLSX format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
