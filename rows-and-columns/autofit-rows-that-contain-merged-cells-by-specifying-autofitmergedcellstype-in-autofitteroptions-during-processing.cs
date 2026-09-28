// Title: How to auto‑fit rows with merged cells using AutoFitMergedCellsType in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that merges a cell range, enables text wrapping, and calls AutoFitRows with AutoFitterOptions set to AutoFitMergedCellsType.EachLine and AutoFitWrappedTextType.Paragraph. | Show an example of configuring AutoFitterOptions to expand each row in a merged area while preserving paragraph wrapping in Aspose.Cells. | Create a workbook that demonstrates adjusting row height for a merged A1:B3 range using AutoFitRows and the appropriate AutoFitMergedCellsType.
// Common Searches: Aspose.Cells C# auto fit rows for merged cells each line option | How to adjust row height of merged cells with text wrap in Aspose.Cells | Example of auto‑fitting rows containing merged cells in an Excel workbook using Aspose.Cells | C# code to auto‑fit rows when cells are merged and text is wrapped in Aspose.Cells | Set options to expand merged cell rows in Aspose.Cells .NET
// Tags: auto-fit rows for merged cells Aspose.Cells | AutoFitterOptions merged cells option C# | row height adjustment wrapped text Aspose.Cells | Excel merged range auto-fit .NET | configure auto-fit merged cells Aspose.Cells

using System;
using Aspose.Cells;

namespace AutoFitMergedCellsDemo
{
    // The example creates a new workbook, merges cells A1:B3, enables text wrapping on the merged cell, configures AutoFitterOptions with AutoFitMergedCellsType.EachLine and AutoFitWrappedTextType.Paragraph, then calls sheet.AutoFitRows(options) to automatically adjust the height of each row in the merged area before saving the file as AutoFitMergedCellsDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into the top‑left cell
            sheet.Cells["A1"].PutValue("This is a sample text for merged cells auto‑fit demonstration. " +
                                       "It is intentionally long to require row height adjustment when the cells are merged.");

            // Merge a range of cells (A1:B3)
            sheet.Cells.Merge(0, 0, 3, 2); // rows 0‑2, columns 0‑1

            // Enable text wrapping so the content can span multiple lines
            Style style = sheet.Cells["A1"].GetStyle();
            style.IsTextWrapped = true;
            sheet.Cells["A1"].SetStyle(style);

            // Configure AutoFitterOptions to auto‑fit merged cells.
            // AutoFitMergedCellsType.EachLine expands the height of every row in the merged area.
            AutoFitterOptions options = new AutoFitterOptions
            {
                AutoFitMergedCellsType = AutoFitMergedCellsType.EachLine,
                AutoFitWrappedTextType = AutoFitWrappedTextType.Paragraph
            };

            // Auto‑fit all rows in the worksheet using the specified options
            sheet.AutoFitRows(options);

            // Save the workbook to a file
            workbook.Save("AutoFitMergedCellsDemo.xlsx");
        }
    }
}
