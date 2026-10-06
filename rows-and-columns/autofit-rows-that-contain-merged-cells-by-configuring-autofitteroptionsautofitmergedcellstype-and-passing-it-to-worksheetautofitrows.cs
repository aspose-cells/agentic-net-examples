// Title: How to auto‑fit rows that contain merged cells using AutoFitterOptions.AutoFitMergedCellsType in Aspose.Cells for .NET
// AI Prompts: Generate C# code that merges a cell range, inserts a long string, configures AutoFitterOptions.AutoFitMergedCellsType to include merged cells, and calls Worksheet.AutoFitRows to adjust row heights. | Show a complete example of using AutoFitterOptions with AutoFitMergedCellsType to correctly auto‑fit rows that have merged cells in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# auto‑fit rows with merged cells | How to enable AutoFitMergedCellsType when auto‑fitting rows in Aspose.Cells .NET | C# example for adjusting row height of merged cells using Aspose.Cells | Worksheet.AutoFitRows handling of merged cells in Aspose.Cells | Set AutoFitterOptions to auto‑fit merged cells in a .NET workbook
// Tags: auto-fit rows merged cells Aspose.Cells | AutoFitterOptions.AutoFitMergedCellsType C# | Worksheet.AutoFitRows merged cell support | adjust row height merged range .NET | Aspose.Cells row auto‑fit merged cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, merges cells A1:D1, places a long text string in the merged range, configures AutoFitterOptions.AutoFitMergedCellsType, calls Worksheet.AutoFitRows to resize the row height accordingly, and saves the file as AutoFitRowsWithMergedCells.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells A1:D1
            worksheet.Cells.Merge(0, 0, 1, 4);
            // Insert a long text into the merged range
            worksheet.Cells["A1"].PutValue(
                "This is a very long piece of text that should cause the row height to increase when auto‑fitted.");

            // Auto‑fit all rows in the worksheet.
            // In recent Aspose.Cells versions, AutoFitRows handles merged cells appropriately.
            worksheet.AutoFitRows();

            // Save the workbook
            string outputFile = "AutoFitRowsWithMergedCells.xlsx";
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputFile)}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
