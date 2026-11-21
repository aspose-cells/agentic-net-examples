// Title: How to split an Excel worksheet view vertically at column 5 using Aspose.Cells SplitPanes in C#
// Common Searches: Aspose.Cells C# split worksheet view at column F without freezing panes | Set vertical split pane at column index 5 using SplitPanes in .NET | Programmatically create a vertical split in an Excel sheet with Aspose.Cells | C# code to split Excel worksheet view vertically at a specific column | Using SplitPanes.SplitColumn to separate worksheet view in Aspose.Cells
// Tags: SplitPanes vertical split column Aspose.Cells | SplitColumn property C# Aspose.Cells | Excel worksheet view split Aspose.Cells | Create split pane without freeze Aspose.Cells | Save workbook after splitting view Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, configures SplitPanes to split the view vertically at column index 5 (column F) with no row split and Freeze set to false, ensures the output directory exists, saves the workbook as SplitPaneExample.xlsx, and outputs a success message.
class SplitWorksheetExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet using dynamic to avoid compile‑time binding issues
            dynamic sheet = workbook.Worksheets[0];

            // Split the view vertically at column 5 (zero‑based index, column F)
            sheet.SplitPanes.SplitColumn = 5;   // column where the split occurs
            sheet.SplitPanes.SplitRow = 0;      // no row split for a vertical split
            sheet.SplitPanes.Freeze = false;    // ensure the split is not frozen

            // Define output file path
            string outputPath = "SplitPaneExample.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
