using Aspose.Cells;
using System;

class SetRowHeightWithAutoFitDemo
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data that will affect row height
        cells["A1"].PutValue("This is a very long text that should cause the row to expand when auto‑fitted.");
        // Enable text wrapping so the row may need to grow
        Style style = cells["A1"].GetStyle();
        style.IsTextWrapped = true;
        cells["A1"].SetStyle(style);

        // Define the minimum row height (in points)
        double minHeight = 30.0;

        // Set the minimum height before auto‑fitting (SetRowHeight rule)
        cells.SetRowHeight(0, minHeight);

        // Auto‑fit the row based on its content (AutoFitRow rule)
        worksheet.AutoFitRow(0);

        // After auto‑fit, ensure the height is not less than the minimum
        double actualHeight = cells.GetRowHeight(0);
        if (actualHeight < minHeight)
        {
            cells.SetRowHeight(0, minHeight);
        }

        // Save the workbook (save rule)
        workbook.Save("SetRowHeightWithAutoFit.xlsx");
    }
}