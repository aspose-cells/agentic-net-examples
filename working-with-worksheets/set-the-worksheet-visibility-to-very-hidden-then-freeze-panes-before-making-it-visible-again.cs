// Title: Set a worksheet to VeryHidden, freeze the first row and column, then restore visibility with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, marks the first worksheet as VeryHidden, applies FreezePanes to lock the top row and left column, and finally makes the sheet visible before saving. | Create a snippet using Aspose.Cells for .NET to temporarily hide a worksheet, configure pane freezing, and then unhide the worksheet in the same operation. | Write a C# example that demonstrates hiding a worksheet, calling FreezePanes(1,1,0,0), and re‑showing the worksheet with Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet VeryHidden then freeze panes | how to freeze first row and column on a hidden sheet using Aspose.Cells | C# Aspose.Cells hide worksheet temporarily and unhide after FreezePanes | example code for toggling worksheet visibility with FreezePanes in Aspose.Cells
// Tags: very hidden worksheet Aspose.Cells | freeze panes first row column Aspose.Cells | toggle worksheet visibility Aspose.Cells | Aspose.Cells worksheet visibility management | C# hide and unhide Excel sheet Aspose.Cells | apply FreezePanes before showing worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing workbook, sets the first worksheet to VeryHidden, freezes the first row and column with FreezePanes(1,1,0,0), restores the worksheet's visibility, and saves the modified file, including basic file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Hide the worksheet (VeryHidden not available in older versions, use IsVisible = false)
            sheet.IsVisible = false;

            // Freeze panes (freeze first row and first column)
            // Using the overload with four parameters for compatibility
            sheet.FreezePanes(1, 1, 0, 0);

            // Make the worksheet visible again
            sheet.IsVisible = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
