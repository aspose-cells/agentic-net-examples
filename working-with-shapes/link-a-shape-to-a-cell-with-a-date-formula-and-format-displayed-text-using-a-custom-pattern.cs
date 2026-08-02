// Title: Link a label shape to a date cell with a custom format using Aspose.Cells for .NET
// Description: Creates a workbook, writes the current date to A1, applies the custom pattern "dd-mmm-yyyy", adds a label shape, links the shape to the formatted cell via SetLinkedCell (A1 style, locale‑aware), and saves the file as LinkedShapeDate.xlsx.
// Keywords: Aspose.Cells | C# | SetLinkedCell | label shape | custom date format | Excel shape binding | locale aware linking | dynamic date display
// Common Searches: Aspose.Cells link shape to cell C# | SetLinkedCell custom date format | label shape display cell value Aspose | bind Excel shape to date cell .NET | locale aware SetLinkedCell example
// Developer Intent: Bind a label shape to a cell that holds a date and show the date using a user‑defined format.
// Use Cases: Generate reports where a shape always shows the current date in a specific pattern. | Create invoices with a linked shape that displays the due date and updates automatically. | Design dashboards where shapes reflect cell values formatted as custom dates.
// AI Prompts: Write C# code with Aspose.Cells to link a label shape to a date cell and apply a custom date format. | Explain the SetLinkedCell parameters for reference style and locale when linking a shape to a formatted date. | Show how to ensure a linked shape refreshes when the underlying date cell changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes the current date to A1, applies the custom pattern "dd-mmm-yyyy", adds a label shape, links the shape to the formatted cell via SetLinkedCell (A1 style, locale‑aware), and saves the file as LinkedShapeDate.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a date value into cell A1
        Cell dateCell = sheet.Cells["A1"];
        dateCell.PutValue(DateTime.Now);

        // Apply a custom date format (e.g., "dd-mmm-yyyy") to the cell
        Style dateStyle = workbook.CreateStyle();
        dateStyle.Custom = "dd-mmm-yyyy";
        dateCell.SetStyle(dateStyle);

        // Add a label shape that will display the linked cell value
        // Parameters: upper left row, upper left column, top offset, left offset, width, height (in pixels)
        Label label = sheet.Shapes.AddLabel(2, 2, 5, 5, 150, 30);

        // Link the shape to the cell containing the date
        // formula: "$A$1", isR1C1 = false (A1 style), isLocal = true (locale‑aware)
        label.SetLinkedCell("$A$1", false, true);

        // Save the workbook
        workbook.Save("LinkedShapeDate.xlsx");
    }
}
