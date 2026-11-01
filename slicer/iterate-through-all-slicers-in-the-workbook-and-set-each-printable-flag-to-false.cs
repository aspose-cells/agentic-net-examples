// Title: How to make every slicer in an Excel workbook non‑printable with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that opens a .xlsx file, iterates through all worksheets and their slicers, sets each slicer's IsPrintable property to false, and saves the workbook. | Create a reusable C# method that receives a Workbook object and disables printing for all slicers it contains via Aspose.Cells. | Write a .NET console application that loads an existing Excel workbook, marks every slicer as non‑printable, and writes the modified file to a new location.
// Common Searches: Aspose.Cells C# set slicer IsPrintable false for all worksheets | disable printing of slicers in Excel using Aspose.Cells .NET | iterate through workbook slicers and change printable flag Aspose.Cells | C# code to make slicers non‑printable when saving Excel file with Aspose
// Tags: Aspose.Cells set slicer IsPrintable | C# disable slicer printing Excel | iterate workbook slicers Aspose.Cells | non‑printable slicers .xlsx Aspose | Excel slicer printable property .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The example loads an existing Excel workbook (or creates a new one), loops over each worksheet and every slicer on those sheets, sets the slicer's IsPrintable flag to false, ensures the output directory exists, and saves the updated workbook to the specified path.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Load existing workbook or create a new one if the file is missing
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all slicers on the worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Make slicer non‑printable
                    slicer.IsPrintable = false;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
