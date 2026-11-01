// Title: C# example: Remove unwanted worksheets by index list using Aspose.Cells and save a filtered workbook
// Common Searches: how to keep only specific worksheets by index when processing an Excel file with Aspose.Cells in C# | C# Aspose.Cells remove sheets not in a given index array | filter Excel worksheets using a list of zero‑based indices before saving with Aspose.Cells | programmatically prune Excel workbook sheets by index using Aspose.Cells .NET | Aspose.Cells load workbook then delete sheets based on user supplied indices C#
// Tags: Aspose.Cells worksheet removal by index C# | filter Excel sheets using index list Aspose.Cells | selective worksheet pruning .NET Excel | load workbook and delete sheets Aspose.Cells | C# Excel workbook sheet filtering Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample verifies the source file, defines a zero‑based list of worksheet indices to retain, loads the workbook with Aspose.Cells, iterates the worksheets in reverse order removing any sheet whose index is not in the list, and saves the resulting filtered workbook to a new file while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // User‑provided list of worksheet indices to keep (0‑based)
            List<int> sheetsToKeep = new List<int> { 0, 2, 4 };

            // Load the workbook normally
            Workbook workbook = new Workbook(inputPath);

            // Remove worksheets whose indices are NOT in the allowed list
            // Iterate in reverse order to avoid index shifting when removing sheets
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (!sheetsToKeep.Contains(i))
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the filtered workbook to a new file
            const string outputPath = "filtered_output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
