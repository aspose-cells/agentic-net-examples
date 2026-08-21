// Title: C# – Add a label shape linked to a custom‑formatted date cell with Aspose.Cells
// Description: Creates a new workbook, writes the current DateTime to cell A1, applies the custom number format "dd-mmm-yyyy", inserts a label shape, links the shape to the cell using SetLinkedCell, refreshes the shape to show the formatted date, and saves the file as ShapeLinkedDate.xlsx.
// Keywords: Aspose.Cells C# | link shape to cell | custom date format Excel | SetLinkedCell method | label shape update | UpdateSelectedValue | Excel shape automation | dynamic date label
// Common Searches: Aspose.Cells link label shape to cell C# | display formatted date in Excel shape using Aspose | SetLinkedCell isR1C1 false example | how to refresh linked shape value Aspose.Cells | add label shape with date in .NET
// Developer Intent: Link a label shape to a worksheet cell that contains a date formatted with a custom pattern and have the shape display that formatted date automatically.
// Use Cases: Generate a report header that always shows the generation date inside a shape. | Create an invoice template where the invoice date appears in a label shape and updates when the source cell changes. | Build a dashboard with dynamic date labels that reflect the latest data without manual edits.
// AI Prompts: Write C# code with Aspose.Cells to add a label shape linked to cell A1 formatted as "dd-mmm-yyyy" and ensure the shape displays the formatted date. | Explain the role of the isR1C1 and isLocal parameters in SetLinkedCell and how they influence cell linking. | Suggest alternative methods to refresh a linked shape after the source cell value is modified in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, writes the current DateTime to cell A1, applies the custom number format "dd-mmm-yyyy", inserts a label shape, links the shape to the cell using SetLinkedCell, refreshes the shape to show the formatted date, and saves the file as ShapeLinkedDate.xlsx.
class ShapeLinkedDateExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a DateTime value into cell A1
        Cell dateCell = worksheet.Cells["A1"];
        dateCell.PutValue(DateTime.Now);

        // Apply a custom date format to the cell (e.g., "dd-mmm-yyyy")
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-mmm-yyyy";
        dateCell.SetStyle(dateStyle);

        // Add a label shape that will display the linked cell value
        // Parameters: upper left row, upper left column, upper left offset X, upper left offset Y, width, height
        Label label = (Label)worksheet.Shapes.AddLabel(2, 2, 0, 0, 120, 30);

        // Link the shape to cell A1 using the SetLinkedCell method
        // isR1C1 = false (A1 style), isLocal = true (locale‑aware)
        label.SetLinkedCell("$A$1", false, true);

        // Refresh the shape so it shows the current value of the linked cell
        label.UpdateSelectedValue();

        // Save the workbook to a file
        workbook.Save("ShapeLinkedDate.xlsx");
    }
}
