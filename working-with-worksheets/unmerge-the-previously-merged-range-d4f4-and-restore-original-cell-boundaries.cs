// Title: Programmatically unmerge the cell range D4:F4 in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# code to unmerge a specific range D4:F4 with Aspose.Cells | Aspose.Cells .NET how to split merged cells D4 to F4 | Unmerge merged cells in Excel using Aspose.Cells C# example | Restore original cell boundaries after merging with Aspose.Cells | Aspose.Cells UnMerge method usage for Excel range D4:F4
// Tags: Aspose.Cells UnMerge method C# | unmerge cell range Excel Aspose.Cells | remove merged cells D4:F4 Aspose.Cells | restore original cell boundaries Aspose.Cells .NET | Excel worksheet cell unmerge operation C# | programmatic unmerge merged range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads an existing workbook (or creates one with cells D4:F4 merged), obtains the D4:F4 range, checks if it is merged, calls the UnMerge() method to split the cells back to their original boundaries, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Paths for the input and output Excel files
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook;

        // Load an existing workbook if it exists; otherwise create a sample workbook with the merged range D4:F4
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load '{inputPath}': {ex.Message}");
                return;
            }
        }
        else
        {
            // Create a new workbook and merge D4:F4 to simulate the pre‑condition
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Merge cells D4 to F4
            AsposeRange sampleRange = ws.Cells.CreateRange("D4", "F4");
            sampleRange.Merge();
            ws.Cells["D4"].PutValue("Merged Sample");

            try
            {
                workbook.Save(inputPath);
                Console.WriteLine($"Sample workbook created and saved as '{inputPath}' with merged range D4:F4.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save sample workbook: {ex.Message}");
                return;
            }
        }

        // Work on the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Obtain the range representing D4:F4
        AsposeRange targetRange = worksheet.Cells.CreateRange("D4", "F4");

        // Verify that the range is currently merged before attempting to unmerge
        if (worksheet.Cells["D4"].IsMerged)
        {
            try
            {
                // Unmerge the range, restoring the original individual cells
                targetRange.UnMerge();
                Console.WriteLine("Successfully unmerged the range D4:F4.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during UnMerge operation: {ex.Message}");
                return;
            }
        }
        else
        {
            Console.WriteLine("The range D4:F4 is not merged; no unmerge operation performed.");
        }

        // Save the modified workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Modified workbook saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save the modified workbook: {ex.Message}");
        }
    }
}
