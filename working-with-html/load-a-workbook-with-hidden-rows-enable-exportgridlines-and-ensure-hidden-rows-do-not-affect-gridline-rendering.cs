// Title: Export an Excel worksheet to PDF with gridlines visible while hidden rows remain excluded using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export PDF with gridlines but ignore hidden rows | How to keep hidden rows hidden when saving Excel as PDF with gridlines in .NET | Enable PrintGridlines and hide rows during PDF conversion using Aspose.Cells | Aspose.Cells PDF export preserving worksheet gridlines without showing hidden rows | Export Excel to PDF with visible gridlines and hidden rows excluded in C#
// Tags: Aspose.Cells PDF export with gridlines | PrintGridlines property C# | exclude hidden rows during PDF conversion | Excel to PDF gridline rendering Aspose.Cells | worksheet IsGridlinesVisible Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, sets both IsGridlinesVisible and PageSetup.PrintGridlines to true, and saves the workbook as a PDF. Aspose.Cells automatically omits hidden rows while rendering gridlines, producing a PDF where gridlines are visible but hidden rows do not affect the layout.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook that contains hidden rows/columns
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Enable gridlines for export (both view and when saving to PDF/Image)
            sheet.IsGridlinesVisible = true;          // shows gridlines in the UI
            sheet.PageSetup.PrintGridlines = true;    // includes gridlines in exported files

            // Aspose.Cells automatically respects hidden rows and columns when exporting.
            // No need to modify visibility flags; the hidden state (IsHidden) is sufficient.

            // Save the workbook as PDF with gridlines exported and hidden rows/columns ignored
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
