using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Set sample column widths (in character units)
        cells.SetColumnWidth(0, 30); // Column A
        cells.SetColumnWidth(1, 15); // Column B

        // Find the maximum column width in the used range
        double maxWidth = 0;
        int usedColumns = cells.MaxColumn + 1; // MaxColumn is zero‑based
        for (int col = 0; col < usedColumns; col++)
        {
            double width = cells.GetColumnWidth(col);
            if (width > maxWidth)
                maxWidth = width;
        }

        // Determine zoom level based on the widest column
        // Example heuristic: wider columns -> zoom out, narrower columns -> zoom in
        int zoom = maxWidth > 25 ? 80 : 120;

        // Clamp zoom to the allowed range (10% – 400%)
        zoom = Math.Max(10, Math.Min(400, zoom));

        // Apply the calculated zoom to the worksheet
        worksheet.Zoom = zoom;

        // Save the workbook
        workbook.Save("ZoomBasedOnColumnWidth.xlsx");
    }
}