// Title: How to merge cells G1:G3 and I1:I3 across multiple worksheets and apply a solid background style using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create combined range for G1:G3 and I1:I3 on every worksheet | How to apply a solid background color to the same cells across multiple sheets with Aspose.Cells | Batch formatting non‑contiguous cells using CreateUnionRange in a .NET Excel workbook | C# example for merging G1:G3 and I1:I3 ranges across sheets with Aspose.Cells | Set uniform style for a multi‑sheet range in Aspose.Cells for .NET
// Tags: combined cell range batch formatting | multi‑sheet style application Aspose.Cells | solid fill for non‑contiguous cells | C# Aspose.Cells unified range formatting | Excel workbook unified range styling

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsUnionRangeExample
{
    // The example creates a workbook with two worksheets, builds a union address that includes the ranges G1:G3 and I1:I3 on each sheet, creates a UnionRange via CreateUnionRange, applies a solid light‑blue background style to all cells in the union, and saves the file as MergedRangeFormatted.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Rename the default worksheet and add another one
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets.Add("Sheet2");

                // Build the union address that includes the same ranges on each worksheet
                string unionAddress = "";
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Append the ranges for the current worksheet
                    unionAddress += $"{ws.Name}!G1:G3,{ws.Name}!I1:I3,";
                }
                // Remove the trailing comma
                unionAddress = unionAddress.TrimEnd(',');

                // Create the UnionRange across all worksheets (sheetIndex 0 is required by the API)
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange(unionAddress, 0);

                // Apply batch formatting (e.g., set a solid background color)
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.LightBlue;
                style.Pattern = BackgroundType.Solid;
                unionRange.ApplyStyle(style, new StyleFlag { All = true });

                // Save the workbook
                string outputPath = "MergedRangeFormatted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
