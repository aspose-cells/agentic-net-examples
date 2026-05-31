using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RefreshLinkedShapesExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a checkbox shape to the worksheet (linked to a cell)
        // Parameters: upper left row, upper left column, top, left, height, width
        int row = 2, column = 2;
        int top = 5, left = 5, height = 20, width = 20;
        CheckBox checkBox = (CheckBox)sheet.Shapes.AddCheckBox(row, column, top, left, height, width);

        // Link the checkbox to cell C3 (row index 2, column index 2)
        // The linked cell will store the checkbox state (TRUE/FALSE)
        checkBox.SetLinkedCell("C3", false, false);

        // Set an initial value in the linked cell
        sheet.Cells["C3"].PutValue(true);

        // Refresh the linked shape so its visual state matches the cell value
        // This is required before saving to ensure the shape position/value is up‑to‑date
        sheet.Shapes.UpdateSelectedValue();

        // Save the workbook
        string filePath = "LinkedShapesRefresh.xlsx";
        workbook.Save(filePath);

        // After saving, refresh linked shapes again in case any layout changes occurred
        // (e.g., column width/row height adjustments that affect shape positions)
        sheet.Shapes.UpdateSelectedValue();

        // Optionally re‑save to persist any adjustments made after the first save
        workbook.Save("LinkedShapesRefresh_Refreshed.xlsx");
    }
}