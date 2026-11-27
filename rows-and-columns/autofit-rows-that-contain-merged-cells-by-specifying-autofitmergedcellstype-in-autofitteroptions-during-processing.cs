// Title: Auto‑fit row height for merged cells using AutoFitterOptions.AutoFitMergedCellsType in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit row height when cells are merged across columns | How to configure AutoFitterOptions.AutoFitMergedCellsType to adjust merged row height in .NET | Aspose.Cells example of AutoFitRow with merged cells and text wrapping | Enable AutoFitMergedCells for merged rows in an Excel workbook using C# | Adjust row height for merged cells in Aspose.Cells without losing formatting
// Tags: auto fit merged row height Aspose.Cells | AutoFitterOptions AutoFitMergedCellsType C# | worksheet.AutoFitRow merged cells example | text wrapped merged cells row auto‑fit | Excel row height adjustment Aspose.Cells .NET

using System;
using Aspose.Cells;

// Demonstrates how to merge cells, enable text wrapping, configure AutoFitterOptions with AutoFitMergedCells and AutoFitMergedCellsType (EachLine), and call worksheet.AutoFitRow to automatically adjust the height of rows containing merged cells in a .NET Excel workbook using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate cells with sample data
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A2"].PutValue("This is a long piece of text that will be merged across several columns and should cause the row height to increase when auto‑fitted.");
            worksheet.Cells["A3"].PutValue("Normal cell");
            worksheet.Cells["B3"].PutValue("Short");

            // Merge cells A2:C2 (row 2, columns 0‑2)
            worksheet.Cells.Merge(1, 0, 1, 3); // Merge from row index 1, column 0, spanning 1 row and 3 columns

            // Enable text wrapping for the merged cell to allow multi‑line display
            Style mergedStyle = worksheet.Cells["A2"].GetStyle();
            mergedStyle.IsTextWrapped = true;
            worksheet.Cells["A2"].SetStyle(mergedStyle);

            // Display original row heights
            Console.WriteLine($"Original height of row 2 (index 1): {worksheet.Cells.GetRowHeight(1)}");
            Console.WriteLine($"Original height of row 3 (index 2): {worksheet.Cells.GetRowHeight(2)}");

            // Configure AutoFitterOptions to handle merged cells
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCells = true,
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                IgnoreHidden = false,
                OnlyAuto = false
            };

            // Auto‑fit the row that contains the merged cells (row index 1)
            // The range covers columns 0 to 2 (A to C)
            worksheet.AutoFitRow(1, 0, 2, options);

            // Optionally auto‑fit other rows without merged cells
            worksheet.AutoFitRow(2); // Row 3 (index 2)

            // Display new row heights after auto‑fit
            Console.WriteLine($"New height of row 2 (index 1) after AutoFit: {worksheet.Cells.GetRowHeight(1)}");
            Console.WriteLine($"New height of row 3 (index 2) after AutoFit: {worksheet.Cells.GetRowHeight(2)}");

            // Save the workbook
            string outputPath = "AutoFitMergedRowsDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
