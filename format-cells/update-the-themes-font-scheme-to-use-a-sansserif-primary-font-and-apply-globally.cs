// Title: How to set a global sans‑serif font scheme (major and minor) for an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# change workbook default font to Arial | Set major and minor fonts for all cells in an Excel file with Aspose.Cells | Apply a global Calibri style to a worksheet using Aspose.Cells .NET | Programmatically update Excel theme fonts using Aspose.Cells in C# | How to use StyleFlag to apply a font to an entire worksheet with Aspose.Cells
// Tags: set workbook default font Aspose.Cells | apply global sans-serif font Excel .NET | use StyleFlag to apply font to worksheet | change major font to Arial Aspose.Cells | update minor font Calibri across cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, sets the workbook's default (major) font to Arial, creates a Calibri style for the minor font, applies that style to all cells in the first worksheet using a StyleFlag, and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Update the default (major) font to a sans‑serif font (e.g., Arial)
            workbook.DefaultStyle.Font.Name = "Arial";

            // (Optional) Update the default minor font by setting the style for the whole workbook
            // Here we create a style with the desired minor font and apply it to all cells
            Style minorStyle = workbook.CreateStyle();
            minorStyle.Font.Name = "Calibri";
            StyleFlag flag = new StyleFlag { Font = true };
            workbook.Worksheets[0].Cells.ApplyStyle(minorStyle, flag);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
