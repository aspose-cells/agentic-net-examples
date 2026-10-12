// Title: Determine and set worksheet zoom percentage from column width using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads a column's width with Aspose.Cells, calculates a zoom factor so the column fits within 20 characters, clamps the value between 10% and 400%, and assigns it to Worksheet.Zoom. | Create a reusable C# method that accepts a Worksheet and a target column width, determines the appropriate zoom level based on the actual column width, enforces the allowed zoom range, and sets the worksheet's Zoom property using Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet zoom based on column width calculation | Programmatically adjust Excel zoom level according to column size with Aspose.Cells | C# example for limiting worksheet zoom between 10% and 400% in Aspose.Cells | How to compute zoom percentage from column width in Aspose.Cells .NET
// Tags: Aspose.Cells worksheet zoom calculation from column width | C# set Worksheet.Zoom property programmatically | column width to zoom conversion Aspose.Cells | zoom level clamping 10-400 Aspose.Cells | dynamic worksheet zoom adjustment .NET

using Aspose.Cells;
using System;

// The example creates a workbook, sets column A width to 30 characters, reads the column width, computes a proportional zoom level (reducing zoom when the width exceeds 20 characters), clamps the zoom between 10% and 400%, applies the zoom to the worksheet, and saves the file as ZoomBasedOnColumnWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Example: set column A width (in characters)
        sheet.Cells.SetColumnWidth(0, 30); // 30 characters

        // Retrieve the width of column A (in characters)
        double columnWidth = sheet.Cells.GetColumnWidth(0);

        // Determine an appropriate zoom level based on the column width
        // If the column is wider than 20 characters, reduce zoom proportionally
        int zoom;
        if (columnWidth > 20)
        {
            // Calculate zoom as (desired width / actual width) * 100
            zoom = (int)Math.Round(20.0 / columnWidth * 100);
            // Clamp zoom to Aspose.Cells allowed range (10% - 400%)
            zoom = Math.Max(10, Math.Min(zoom, 400));
        }
        else
        {
            zoom = 100; // Default zoom when width is within normal range
        }

        // Apply the calculated zoom level to the worksheet
        sheet.Zoom = zoom;

        // Save the workbook to a file
        workbook.Save("ZoomBasedOnColumnWidth.xlsx");
    }
}
