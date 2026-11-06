// Title: Merge cells B2:C3 on every worksheet and set a light‑gray fill using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge range B2:C3 on all worksheets | apply background color to merged cells in each sheet with Aspose.Cells | loop through worksheets and format a cell range with solid fill in .NET | set light gray fill for merged cells in an Excel workbook using Aspose.Cells | C# code to merge cells and apply style across multiple worksheets
// Tags: merge B2:C3 range across worksheets Aspose.Cells | apply solid background to merged cells Aspose.Cells | loop through workbook worksheets for cell formatting Aspose.Cells | create light gray fill style Aspose.Cells | C# Excel cell merging with Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The example loads an existing workbook, creates a solid light‑gray style, iterates over every worksheet, merges the range B2:C3, applies the style to the merged range, and saves the updated file, including input validation and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Create a style with light gray background
            Style grayStyle = workbook.CreateStyle();
            grayStyle.ForegroundColor = Color.LightGray;
            grayStyle.Pattern = BackgroundType.Solid;

            // Apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Merge cells B2:C3 (zero‑based indices)
                sheet.Cells.Merge(1, 1, 2, 2); // B2 = row 1, column 1; C3 = row 2, column 2

                // Apply the gray background style to the merged range
                Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("B2:C3");
                mergedRange.ApplyStyle(grayStyle, flag);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
