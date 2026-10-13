// Title: Set worksheet StandardWidth and auto‑fit a specific column range while preserving a manually overridden column width using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook, assigns StandardWidth = 20 to the worksheet, manually sets column C width to 30, populates sample data, and then invokes the AutoFitColumns method for columns B‑D, keeping the manual width unchanged. | Write a C# example with Aspose.Cells that shows how to override a single column's width before calling AutoFitColumns on a selected range, ensuring the overridden column retains its size.
// Common Searches: Aspose.Cells C# set StandardWidth then auto fit columns B to D | preserve manually set column width when using AutoFitColumns in .NET | auto fit a range of columns after defining worksheet StandardWidth Aspose.Cells | how to override a column width before calling AutoFitColumns with Aspose.Cells | C# Aspose.Cells column width precedence StandardWidth versus AutoFitColumns
// Tags: Aspose.Cells StandardWidth property C# | AutoFitColumns method column range Aspose.Cells | manual column width override before autofit Aspose.Cells | column width precedence Aspose.Cells .NET | Excel workbook column width customization Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Creates a workbook, sets worksheet StandardWidth to 20 characters, fills sample data, manually sets column C width to 30, auto‑fits columns B‑D, and saves the file as StandardWidthAutoFit.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Set the standard column width (in characters) for all columns
            sheet.Cells.StandardWidth = 20;

            // Fill some sample data
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["C1"].PutValue("Header3");
            sheet.Cells["D1"].PutValue("Header4");

            sheet.Cells["A2"].PutValue("Short");
            sheet.Cells["B2"].PutValue("A very long text that should cause column to expand");
            sheet.Cells["C2"].PutValue("Medium length");
            sheet.Cells["D2"].PutValue("Another long text that will be auto‑fitted");

            // Manually override the width of column C (index 2)
            sheet.Cells.Columns[2].Width = 30;

            // Auto‑fit columns B to D (indexes 1 to 3)
            // startColumn = 1 (B), totalColumns = 3 (B, C, D)
            sheet.AutoFitColumns(1, 3);

            // Define output file path
            string outputPath = "StandardWidthAutoFit.xlsx";

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
