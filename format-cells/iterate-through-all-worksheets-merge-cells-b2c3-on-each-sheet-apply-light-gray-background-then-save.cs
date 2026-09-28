// Title: Iterate all worksheets to merge cells B2:C3 and set a light gray background using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, loops through every worksheet, merges the range B2:C3, and applies a light‑gray solid fill style to the merged cells before saving. | Create a reusable C# method that receives a Workbook object, merges B2:C3 on each sheet, builds a style with a light gray background, applies it to the merged range, and returns the updated workbook.
// Common Searches: how to merge a cell range on every sheet using Aspose.Cells C# | apply a gray background to merged cells across all worksheets in Aspose.Cells | Aspose.Cells loop through worksheets and format B2:C3 range | C# Aspose.Cells set solid fill color for merged cells in multiple sheets | save workbook after applying style to merged range with Aspose.Cells .NET
// Tags: Aspose.Cells API merge cells | Aspose.Cells API set background color | Aspose.Cells workbook worksheet iteration | Aspose.Cells range style application | Aspose.Cells save workbook after formatting

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// // Loads (or creates) an Excel workbook, iterates through each worksheet, merges the B2:C3 range, applies a light gray solid fill style to the merged cells, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Load existing workbook if it exists; otherwise create a new one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook(); // creates a default workbook with one worksheet
            }

            // Define light gray background color
            Color lightGray = Color.FromArgb(211, 211, 211);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Merge cells B2:C3 (row index 1, column index 1, spanning 2 rows and 2 columns)
                sheet.Cells.Merge(1, 1, 2, 2);

                // Create a style with light gray solid fill
                Style style = workbook.CreateStyle();
                style.ForegroundColor = lightGray;
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the merged range
                Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("B2:C3");
                mergedRange.ApplyStyle(style, new StyleFlag() { All = true });
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
