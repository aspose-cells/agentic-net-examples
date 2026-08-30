// Title: Auto‑fit row heights for merged cells in an Aspose.Cells .NET workbook using AutoFitterOptions
// AI Prompts: Write C# code that merges a cell range, enables text wrapping, sets AutoFitterOptions.AutoFitMergedCellsType to EachLine, and calls Worksheet.AutoFitRows to adjust row heights. | Show how to combine AutoFitWrappedTextType.Paragraph with AutoFitMergedCellsType.EachLine in AutoFitterOptions before saving an Excel file with Aspose.Cells. | Provide a step‑by‑step example that creates a workbook, merges A1:B3, applies wrapped‑text style, configures AutoFitterOptions, and auto‑fits rows programmatically.
// Common Searches: Aspose.Cells C# auto fit rows for merged cells each line option | How to adjust row height of merged range using AutoFitterOptions in .NET | AutoFitRows with AutoFitWrappedTextType Paragraph for wrapped text in Aspose.Cells | C# example for merging cells and auto‑fitting row height in Excel workbook
// Tags: AutoFitterOptions row height adjustment | AutoFitMergedCellsType EachLine usage | merged range row height auto‑fit technique | wrapped text AutoFitWrappedTextType Paragraph | Aspose.Cells Excel row height for merged cells

using System;
using Aspose.Cells;

namespace AutoFitMergedCellsDemo
{
    // // Demonstrates creating a workbook, merging cells A1:B3, enabling text wrap, configuring AutoFitterOptions with AutoFitMergedCellsType=EachLine and AutoFitWrappedTextType=Paragraph, auto‑fitting rows, and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell
            worksheet.Cells["A1"].PutValue("This is a sample text for merged cells auto‑fit demonstration. " +
                                          "It contains enough characters to require multiple lines when wrapped.");

            // Merge a range of cells (A1:B3)
            worksheet.Cells.Merge(0, 0, 3, 2);

            // Enable text wrapping so the content can span multiple lines
            Style style = worksheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(style);

            // Configure AutoFitterOptions to consider merged cells (EachLine expands every row in the merged area)
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph // optional, improves wrapped‑text handling
            };

            // Auto‑fit rows using the configured options (feature rule: AutoFitRows(AutoFitterOptions))
            worksheet.AutoFitRows(options);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("AutoFitMergedCellsDemo.xlsx");
        }
    }
}
