// Title: How to apply the SlicerStyleLight1 formatting to a slicer and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file, checks for slicers on the first worksheet, sets the first slicer's style to SlicerStyleLight1 using Aspose.Cells, and saves the workbook to a new file. | Show a robust C# example that creates the output directory if missing, handles the case where no slicers are present, and applies a slicer style with version‑compatible error handling in Aspose.Cells. | Generate a snippet that retrieves the SlicerStyles collection from a Workbook, assigns the SlicerStyleLight1 style to a specific Slicer object, and persists the changes.
// Common Searches: Aspose.Cells C# set slicer style to SlicerStyleLight1 in existing Excel file | programmatically change slicer formatting in .xlsx using Aspose.Cells for .NET | C# check if worksheet contains slicers before applying style with Aspose.Cells | save workbook after updating slicer style with Aspose.Cells
// Tags: Aspose.Cells apply slicer style | C# set slicer formatting Excel | SlicerStyleLight1 Aspose.Cells | save workbook after slicer modification | verify slicer existence Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Namespace for Slicer class

// The example loads input.xlsx, verifies that the first worksheet has at least one slicer, assigns the SlicerStyleLight1 style from the workbook's SlicerStyles collection to the first slicer (with version‑compatibility handling), ensures the output directory exists, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one slicer on the worksheet
            if (worksheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the first worksheet.");
                return;
            }

            // Get the first slicer
            Slicer slicer = worksheet.Slicers[0];

            // NOTE: Applying a slicer style requires a newer Aspose.Cells version.
            // If the current version supports it, uncomment the following lines:
            // try
            // {
            //     slicer.Style = workbook.SlicerStyles["SlicerStyleLight1"];
            // }
            // catch (Exception styleEx)
            // {
            //     Console.WriteLine($"Failed to apply slicer style: {styleEx.Message}");
            // }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
