// Title: Add a rectangle shape that displays a formatted date from a cell using the TEXT function in Aspose.Cells for .NET
// AI Prompts: Insert a rectangle shape on a worksheet and assign its Text property to =TEXT(A1,"dd-MMM-yyyy") with Aspose.Cells. | Apply a built‑in date number format to a cell, then bind a shape’s displayed text to that cell using a custom date pattern via the TEXT function. | After setting a TEXT formula on a shape, call workbook.CalculateFormula() and read the evaluated shape text.
// Common Searches: Aspose.Cells bind shape text to a cell using TEXT formula | display a custom formatted date inside a shape in Aspose.Cells .NET | link rectangle shape to cell A1 with dd-MMM-yyyy format Aspose.Cells | recalculate formulas to refresh shape text in an Aspose.Cells workbook
// Tags: add rectangle shape Aspose.Cells | shape text formula TEXT Aspose.Cells | custom date format cell Aspose.Cells | calculate formulas shape text .NET | save workbook with linked shape Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a new workbook, writes a DateTime value to cell A1, applies the built‑in date format (ID 14), adds a rectangle shape, sets its Text property to a TEXT formula that formats the date as "dd-MMM-yyyy", recalculates all formulas so the shape shows the formatted date, reads the evaluated text, and saves the file as ShapeLinkedToCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a date value into cell A1
            DateTime sampleDate = new DateTime(2023, 12, 25);
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue(sampleDate);

            // Apply a built‑in date format (Number format ID 14)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Number = 14;
            cell.SetStyle(dateStyle);

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 100, 30);

            // Link the shape's text to cell A1 using the TEXT function with a custom date format
            shape.Text = "=TEXT(A1,\"dd-MMM-yyyy\")";

            // Recalculate all formulas so the shape text is updated
            workbook.CalculateFormula();

            // Verify: read the displayed text from the shape
            string shapeText = shape.Text; // After calculation this holds the evaluated string
            Console.WriteLine("Shape text: " + shapeText); // Expected: 25-Dec-2023

            // Save the workbook (optional, demonstrates that the shape is persisted)
            workbook.Save("ShapeLinkedToCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
