// Title: Load an Excel workbook with Aspose.Cells in C# and keep only defined names that contain "Total"
// Common Searches: Aspose.Cells C# filter defined names that contain the word Total when opening a workbook | How to delete named ranges that do not include 'Total' from an Excel file using Aspose.Cells | Check that only specific defined names remain after loading an Excel workbook with Aspose.Cells | C# example for retaining only defined names matching a keyword in Aspose.Cells | Remove unwanted defined names during workbook load with Aspose.Cells API
// Tags: Aspose.Cells defined name filtering | retain named ranges with keyword | remove non-matching defined names | validate defined name list after load | C# Excel defined name manipulation

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The code loads an existing Excel file (or creates a new one), iterates through the workbook's defined names, removes any name that does not contain the word "Total" (case‑insensitive), verifies that only the desired names remain, and saves the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new empty workbook.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // empty workbook as fallback
            }

            // Collect defined names that do NOT contain "Total"
            List<string> namesToRemove = new List<string>();
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                if (!definedName.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    namesToRemove.Add(definedName.Text);
                }
            }

            // Remove the unwanted defined names from the workbook
            foreach (string nameText in namesToRemove)
            {
                // NameCollection.Remove returns true if the name existed and was removed
                workbook.Worksheets.Names.Remove(nameText);
            }

            // Verify that only names containing "Total" remain
            foreach (Name definedName in workbook.Worksheets.Names)
            {
                if (!definedName.Text.Contains("Total", StringComparison.OrdinalIgnoreCase))
                {
                    throw new Exception($"Unexpected defined name retained: {definedName.Text}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional, to see the result)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
