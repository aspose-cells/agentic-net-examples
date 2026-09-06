// Title: How to merge cells Z4:AA4 and set a 45° text rotation in an XLSX workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Merge the range Z4:AA4, create a style with a 45‑degree rotation, apply it via a StyleFlag, and save the workbook as XLSX using Aspose.Cells in C#. | Load an existing XLSX file, combine cells Z4 and AA4, define a custom rotation style, assign the style to the merged range, and export the result with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge cells Z4 AA4 and rotate text 45 degrees | set text rotation for a merged range using Aspose.Cells .NET | how to apply a 45 degree text orientation to specific cells in Excel with Aspose.Cells | C# code to merge cells and apply rotation style before saving as XLSX
// Tags: merge cells Z4 AA4 Aspose.Cells | text rotation style Aspose.Cells C# | StyleFlag rotation attribute Aspose.Cells | save workbook as XLSX Aspose.Cells | custom rotation style creation Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads input.xlsx, merges cells Z4:AA4 on the first worksheet, creates a style with a 45-degree text rotation, applies the style to the merged range using a StyleFlag, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells Z4:AA4 (row index 3, column index 25, 1 row, 2 columns)
            sheet.Cells.Merge(3, 25, 1, 2);

            // Create a style with 45° text rotation
            Style rotationStyle = workbook.CreateStyle();
            rotationStyle.RotationAngle = 45;

            // Specify that only rotation should be applied
            StyleFlag flag = new StyleFlag();
            flag.Rotation = true;

            // Apply the style to the merged range
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("Z4:AA4");
            mergedRange.ApplyStyle(rotationStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
