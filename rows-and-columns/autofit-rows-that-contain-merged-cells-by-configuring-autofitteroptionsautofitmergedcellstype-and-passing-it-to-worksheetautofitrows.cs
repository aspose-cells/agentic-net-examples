// Title: Auto‑fit rows that contain merged cells in Aspose.Cells for .NET using AutoFitterOptions.AutoFitMergedCellsType (C#)
// Common Searches: C# Aspose.Cells auto fit rows that have merged cells using AutoFitterOptions | How to preserve merged cell height when auto‑fitting rows in Aspose.Cells .NET | Set AutoFitMergedCellsType option for Worksheet.AutoFitRows in C# | Aspose.Cells example: auto‑fit row height after merging cells | Row height adjustment for merged ranges with AutoFitRows in Aspose.Cells
// Tags: AutoFitterOptions.AutoFitMergedCellsType C# | Worksheet.AutoFitRows merged cells handling | Aspose.Cells row height auto‑fit merged range | .NET workbook auto‑fit rows merged cells | C# merge cells then auto‑fit rows Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, merges cells A1:D1 on the first worksheet, configures AutoFitterOptions.AutoFitMergedCellsType to consider merged cells, calls Worksheet.AutoFitRows to adjust row heights accordingly, and saves the result as AutoFitMergedRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("This is a merged cell");
            sheet.Cells["B1"].PutValue("Additional text");

            // Merge cells A1:D1 (row 0, columns 0-3)
            sheet.Cells.Merge(0, 0, 1, 4);

            // Auto‑fit all rows in the worksheet.
            // The default behavior considers merged cells for row autofit.
            sheet.AutoFitRows();

            // Define output file path
            string outputPath = "AutoFitMergedRows.xlsx";

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
