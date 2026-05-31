using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ShapeLinkedCellPercentageDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a numeric value into cell A1 (e.g., 0.25)
        Cell linkedCell = worksheet.Cells["A1"];
        linkedCell.PutValue(0.25);

        // Apply a custom number format that displays the value as a percentage using TEXT
        // Here we use the built‑in percentage format (Number = 10) which sets IsPercent = true
        Style style = linkedCell.GetStyle();
        style.Number = 10; // Built‑in percentage format
        linkedCell.SetStyle(style);

        // Add a label shape to the worksheet
        // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
        Label label = worksheet.Shapes.AddLabel(2, 0, 5, 5, 100, 30);
        label.Text = "Linked Percentage";

        // Link the shape to cell A1
        // formula: "$A$1", isR1C1 = false, isLocal = true (locale‑aware)
        label.SetLinkedCell("$A$1", false, true);

        // Verify that the linked cell's style reports IsPercent = true
        Style verifyStyle = linkedCell.GetStyle();
        Console.WriteLine("IsPercent after linking: " + verifyStyle.IsPercent); // Expected: True

        // Save the workbook
        workbook.Save("ShapeLinkedCellPercentageDemo.xlsx");
    }
}