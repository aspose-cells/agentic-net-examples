// Title: Use Aspose.Cells .NET LoadFilter to load only worksheets with specified zero‑based indices
// AI Prompts: Write C# code that implements an ILoadFilter to load an Excel workbook with Aspose.Cells, keeping only the worksheets whose indices are listed in a user‑provided collection. | Generate a method that accepts a list of worksheet indices, applies a LoadFilter when opening the file, and returns a Workbook containing only those sheets. | Create an example that demonstrates removing unwanted sheets after loading a workbook by iterating the Worksheets collection in reverse based on a given index list.
// Common Searches: aspocells loadfilter keep specific worksheets by index c# | how to open an Excel file with only selected sheets using Aspose.Cells .NET | c# filter worksheets during workbook load based on a list of indices | load only certain sheets from large Excel workbook Aspose.Cells performance
// Tags: loadfilter select worksheets by index Aspose.Cells | c# load specific Excel sheets .NET | aspocells workbook load selected worksheets | filter Excel worksheets during load using index list | remove unwanted sheets after workbook load c#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample defines a zero‑based list of worksheet indices to retain, verifies that the source Excel file exists, loads the workbook with Aspose.Cells, iterates the Worksheets collection in reverse to delete any sheet whose index is not in the list, saves the filtered workbook to a new file, and prints status or error messages.
class Program
{
    static void Main()
    {
        try
        {
            // List of worksheet indices to keep (0‑based)
            List<int> worksheetsToKeep = new List<int> { 0, 2, 4 };

            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook without any filter
            Workbook workbook = new Workbook(inputPath);

            // Remove worksheets whose indices are not in the keep list
            // Iterate in reverse to avoid index shifting issues
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (!worksheetsToKeep.Contains(i))
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
