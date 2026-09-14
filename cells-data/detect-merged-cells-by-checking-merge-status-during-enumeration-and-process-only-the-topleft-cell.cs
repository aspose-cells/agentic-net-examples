// Title: How to detect merged cells and style only the top‑left cell of each merged range using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that iterates over all worksheet cells, checks the IsMerged flag, and applies a yellow background exclusively to the first cell of each merged area. | Show an example that creates merged regions, retrieves the merged range via GetMergedRange, and updates the style of the leading cell while ignoring all inner cells. | Write a script that enumerates worksheet data, prints values of normal cells, and highlights the top‑left cell of every merged block in an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# iterate worksheet and style only the first cell of merged areas | how to change background of top left merged cell in Excel with Aspose.Cells | skip inner cells when processing merged ranges using Aspose.Cells .NET | retrieve merged range coordinates IsMerged Aspose.Cells example | apply style to merged range start cell without affecting other cells Aspose.Cells
// Tags: style top-left merged cell Aspose.Cells | enumerate cells with IsMerged check C# | GetMergedRange usage Aspose.Cells | apply background color to merged start cell .NET | skip inner merged cells Aspose.Cells enumeration

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, merges specific ranges, then loops through all cells, uses the IsMerged property to identify merged areas, processes only the top‑left cell of each merged range by printing its value and setting a yellow background, handles regular cells separately, and finally saves the workbook.
class DetectMergedCellsDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample merged cells for demonstration
            cells.Merge(0, 0, 2, 2); // Merge A1:B2
            cells[0, 0].PutValue("TopLeft");
            cells[1, 1].PutValue("Inner");

            cells.Merge(3, 1, 3, 2); // Merge B4:C6
            cells[3, 1].PutValue("AnotherTop");

            // Enumerate all cells that contain data
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    if (cell.IsMerged)
                    {
                        // Process only the top‑left cell of the merged area
                        AsposeRange mergedRange = cell.GetMergedRange();
                        if (mergedRange != null &&
                            row == mergedRange.FirstRow && col == mergedRange.FirstColumn)
                        {
                            Console.WriteLine($"Top‑left merged cell {cell.Name} value: {cell.Value}");

                            // Example processing: change background color
                            Style style = cell.GetStyle();
                            style.ForegroundColor = Color.Yellow;
                            style.Pattern = BackgroundType.Solid;
                            cell.SetStyle(style);
                        }
                    }
                    else
                    {
                        // Process non‑merged cells as needed
                        if (cell.Value != null)
                        {
                            Console.WriteLine($"Normal cell {cell.Name} value: {cell.Value}");
                        }
                    }
                }
            }

            // Save the workbook
            string outputPath = "MergedCellsProcessed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
