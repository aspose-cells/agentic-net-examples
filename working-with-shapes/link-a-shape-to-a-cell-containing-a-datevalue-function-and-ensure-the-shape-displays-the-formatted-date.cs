// Title: C# – Link a rectangle shape to a DATEVALUE cell and display the formatted date with Aspose.Cells
// Description: Shows how to place a DATEVALUE formula in a cell, apply a date number format, add a rectangle shape, link the shape to the cell using the LinkedCell property, refresh the shape with UpdateSelectedValue, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | LinkedCell | UpdateSelectedValue | shape to cell | DATEVALUE formula | date formatting | rectangle shape | Excel automation
// Common Searches: link shape to cell Aspose.Cells C# | display DATEVALUE result in a shape using Aspose.Cells | update linked shape after changing cell format Aspose.Cells | how to bind a rectangle to a formula cell in Aspose.Cells | Aspose.Cells shape shows formatted date
// Developer Intent: The developer needs to bind a shape to a cell that contains a DATEVALUE formula and ensure the shape shows the cell’s formatted date value.
// Use Cases: Dashboard widgets that reflect calculated dates from formulas. | Automated reports where shapes display due dates derived from DATEVALUE. | Timeline graphics with shapes synchronized to formatted date cells.
// AI Prompts: Provide C# code to link a rectangle shape to a DATEVALUE cell and display the formatted date with Aspose.Cells. | Show how to refresh a shape after changing a cell’s date format using LinkedCell and UpdateSelectedValue in Aspose.Cells. | Explain the steps to bind a shape to a formula cell and keep its displayed value synchronized in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to place a DATEVALUE formula in a cell, apply a date number format, add a rectangle shape, link the shape to the cell using the LinkedCell property, refresh the shape with UpdateSelectedValue, and save the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define a cell that contains a DATEVALUE formula
        Cell dateCell = sheet.Cells["B2"];
        dateCell.Formula = "=DATEVALUE(\"2023-08-15\")";

        // Apply a date number format to the cell (e.g., mm/dd/yyyy)
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Number = 14; // Built‑in date format
        dateCell.SetStyle(dateStyle);

        // Add a rectangle shape that will display the linked cell's value
        // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
        Shape shape = sheet.Shapes.AddRectangle(2, 1, 0, 0, 30, 150);

        // Link the shape to the cell containing the DATEVALUE formula
        shape.LinkedCell = "$B$2";

        // Ensure the shape reflects the current cell value
        shape.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("LinkedShapeDate.xlsx");
    }
}
