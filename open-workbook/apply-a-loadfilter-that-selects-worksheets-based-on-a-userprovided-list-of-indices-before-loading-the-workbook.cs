// Title: Load an Excel workbook with Aspose.Cells in C# and keep only selected worksheets by index
// Common Searches: c# aspnet load excel file with aspose.cells and retain only worksheets at indexes 0 2 4 | how to filter worksheets by index when opening a workbook using Aspose.Cells for .NET | remove unwanted sheets after loading workbook with Aspose.Cells C# example | select specific worksheets during workbook load aspnet aspose.cells | aspose.cells load workbook then delete sheets not in keep list c#
// Tags: load workbook with selected worksheets Aspose.Cells | filter worksheets by index C# | remove sheets not in keep list Aspose.Cells | select worksheets during load Aspose.Cells | worksheet index based filtering Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example checks for the input Excel file, defines a zero‑based list of worksheet indices to keep, loads the workbook with Aspose.Cells, iterates backward through the worksheets removing any whose index is not in the list, and saves the filtered workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Example user‑provided list of worksheet indices to keep (zero‑based)
            List<int> sheetsToKeep = new List<int> { 0, 2, 4 };

            // Load the workbook normally
            Workbook workbook = new Workbook(inputPath);

            // Remove worksheets whose indices are NOT in the keep list
            // Iterate backwards to avoid index shifting when removing sheets
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (!sheetsToKeep.Contains(i))
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
