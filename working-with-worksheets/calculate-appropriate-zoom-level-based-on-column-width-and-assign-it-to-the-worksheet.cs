// Title: Calculate and set worksheet zoom level based on column width using Aspose.Cells for .NET
// Common Searches: calculate worksheet zoom from column width Aspose.Cells C# | set Excel zoom to match column pixel size using Aspose.Cells | programmatically adjust zoom to fit a column in Aspose.Cells .NET | determine zoom percentage for target visible width in Aspose.Cells | how to clamp worksheet zoom range in Aspose.Cells C#
// Tags: zoom property Aspose.Cells C# | GetColumnWidthPixel method Aspose.Cells | column width to zoom conversion Aspose.Cells | zoom range clamping 10-400 Aspose.Cells | fit column to target pixel width Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, sets column A width, retrieves its pixel width with GetColumnWidthPixel, computes a zoom percentage so the column fits within an 800‑pixel target width, clamps the zoom between 10% and 400%, assigns the value to the worksheet's Zoom property, and saves the file as ZoomBasedOnColumnWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Set a specific column width for demonstration
        // Width is measured in characters (default is 8.43)
        sheet.Cells.SetColumnWidth(0, 20); // Column A

        // Get the width of column A in pixels
        int columnWidthPixels = sheet.Cells.GetColumnWidthPixel(0);

        // Desired visible width in pixels (e.g., the width of the window or print area)
        int targetVisibleWidth = 800;

        // Calculate zoom so that the column fits into the target width
        // Zoom (%) = (target width / actual column width) * 100
        int zoom = 100; // default zoom
        if (columnWidthPixels > 0)
        {
            zoom = (int)Math.Round((double)targetVisibleWidth / columnWidthPixels * 100);
        }

        // Clamp zoom to Aspose.Cells allowed range (10% – 400%)
        zoom = Math.Max(10, Math.Min(400, zoom));

        // Apply the calculated zoom level to the worksheet
        sheet.Zoom = zoom;

        // Save the workbook
        workbook.Save("ZoomBasedOnColumnWidth.xlsx");
    }
}
