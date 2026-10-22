// Title: Create a merged title block spanning A1 to F2 with centered bold text using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge a range from A1 to F2 for a report title | How to set centered bold formatting on a merged cell range in Aspose.Cells | Saving a workbook after merging cells in Aspose.Cells .NET | Example of creating a title block that covers multiple rows and columns with Aspose.Cells
// Tags: merge cell range Aspose.Cells C# | centered bold style merged range Aspose.Cells | title block creation Aspose.Cells worksheet | save merged worksheet as XLSX Aspose.Cells | define merged area dimensions Aspose.Cells API

using System;
using Aspose.Cells;

// The sample creates a new workbook, merges cells A1:F2 on the first worksheet to form a title block, inserts "Report Title" with centered bold 16‑point formatting, and saves the file as MergedTitleBlock.xlsx.
class MergeTitleBlock
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells from A1 to F2 to create a title block
        // Cells["A1"] is the top-left cell, Cells["F2"] is the bottom-right cell
        sheet.Cells.Merge(0, 0, 2, 6); // (firstRow, firstColumn, totalRows, totalColumns)

        // Set the title text in the merged cell
        sheet.Cells["A1"].PutValue("Report Title");

        // Optional: Apply some formatting to the title
        Style style = workbook.CreateStyle();
        style.HorizontalAlignment = TextAlignmentType.Center;
        style.VerticalAlignment = TextAlignmentType.Center;
        style.Font.IsBold = true;
        style.Font.Size = 16;
        sheet.Cells["A1"].SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedTitleBlock.xlsx");
    }
}
