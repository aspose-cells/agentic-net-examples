// Title: Set Worksheet Zoom Dynamically from Column Width with Aspose.Cells for .NET
// Description: Demonstrates how to read a column's width, compute a proportional zoom level (100 % up to 20 characters, lower for wider columns), clamp the value between 10 % and 400 %, apply it to the worksheet, and save the workbook as ZoomBasedOnColumnWidth.xlsx.
// Keywords: Aspose.Cells worksheet zoom | column width to zoom conversion | dynamic Excel zoom .NET | set zoom programmatically | zoom range 10 400 percent
// Common Searches: Aspose.Cells set zoom based on column width | calculate Excel zoom from column size .NET | limit worksheet zoom to 10-400 percent Aspose | auto-adjust Excel view using column width
// Developer Intent: Compute an appropriate zoom percentage from a column's width and assign it to the worksheet.
// Use Cases: Fit wide data on screen by reducing zoom when a column exceeds a threshold. | Enforce minimum and maximum zoom levels for consistent report appearance. | Create Excel files where the initial view adapts to the size of a key column.
// AI Prompts: Generate a reusable method that takes a worksheet and column index, calculates the zoom level based on the column width, and applies it while respecting the 10‑400 % limits. | Refactor the sample to support any column, add error handling for invalid widths, and return the applied zoom value. | Explain Aspose.Cells' column width unit (character count) and how it influences the zoom calculation logic.

using System;
using Aspose.Cells;

// Demonstrates how to read a column's width, compute a proportional zoom level (100 % up to 20 characters, lower for wider columns), clamp the value between 10 % and 400 %, apply it to the worksheet, and save the workbook as ZoomBasedOnColumnWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set a specific column width (e.g., column A to 30 characters)
        cells.SetColumnWidth(0, 30);

        // Retrieve the actual column width
        double columnWidth = cells.GetColumnWidth(0);

        // Calculate a zoom level based on the column width.
        // Example logic: keep zoom at 100% for widths up to 20 characters,
        // and reduce proportionally for wider columns.
        int zoom = 100;
        if (columnWidth > 20)
        {
            zoom = (int)(100 * 20 / columnWidth);
        }

        // Ensure the zoom value stays within the allowed range (10% - 400%)
        zoom = Math.Max(10, Math.Min(400, zoom));

        // Apply the calculated zoom to the worksheet
        worksheet.Zoom = zoom;

        // Save the workbook
        workbook.Save("ZoomBasedOnColumnWidth.xlsx");
    }
}
