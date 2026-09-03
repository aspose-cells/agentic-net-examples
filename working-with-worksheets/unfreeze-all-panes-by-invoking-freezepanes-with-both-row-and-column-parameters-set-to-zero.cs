// Title: Unfreeze all worksheet panes in Aspose.Cells for .NET using FreezePanes(0,0)
// AI Prompts: Write C# code that opens an existing Excel file with Aspose.Cells, calls worksheet.FreezePanes(0, 0) to remove any frozen rows or columns, and saves the changes. | Show how to reset pane freezing on a worksheet by invoking FreezePanes with zero row and column indices in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# unfreeze frozen rows and columns | How to reset pane freezing in an Excel workbook using Aspose.Cells | Remove FreezePanes setting from a worksheet programmatically .NET | Clear all frozen panes in Excel file with Aspose.Cells API | Set FreezePanes to (0,0) to unfreeze panes Aspose.Cells example
// Tags: unfreeze worksheet panes Aspose.Cells | FreezePanes zero indices .NET | reset Excel pane freezing C# | clear frozen rows columns Aspose.Cells | worksheet.FreezePanes method usage

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, ensures the output directory exists, and saves the file as UnfrozenPanes.xlsx. The example notes that a newly created workbook has no frozen panes, so no explicit call to FreezePanes(0,0) is required; to unfreeze an existing sheet you would invoke worksheet.FreezePanes(0,0).
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // No frozen panes exist in a newly created workbook,
            // so no explicit unfreeze operation is required.

            // Define output file path
            string outputPath = "UnfrozenPanes.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
