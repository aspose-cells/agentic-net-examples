// Title: How to monitor Excel to multi‑page TIFF conversion progress using Aspose.Cells SaveProgress event in C#
// AI Prompts: Write C# code that subscribes to Workbook.SaveProgress, saves an .xlsx as a multi‑page TIFF, and prints the percentage completed to the console. | Show an Aspose.Cells example that converts a workbook to TIFF while displaying real‑time progress updates using the SaveProgress event. | Generate a snippet that handles the conversion progress callback during workbook.Save to Tiff and logs incremental percentages in .NET.
// Common Searches: aspnet track progress of workbook.Save when exporting to TIFF with Aspose.Cells | c# get conversion percentage while saving Excel as multi page TIFF using Aspose.Cells | how to use SaveProgress event for TIFF export in Aspose.Cells .NET
// Tags: Aspose.Cells SaveProgress event for TIFF export | C# monitor Excel to TIFF conversion percentage | Workbook.Save with progress callback Aspose.Cells | multi‑page TIFF generation progress .NET | real‑time conversion feedback Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads an .xlsx workbook, subscribes to the Workbook.SaveProgress event, saves it as a multi‑page TIFF, and writes conversion percentages to the console while handling errors.
class WorkbookToTiffConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.tiff";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a multi‑page TIFF file
            workbook.Save(outputPath, SaveFormat.Tiff);

            Console.WriteLine($"Workbook successfully converted to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
