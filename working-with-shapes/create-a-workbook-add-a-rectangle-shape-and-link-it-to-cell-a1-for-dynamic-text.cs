// Title: C# – Add a Rectangle Shape Linked to Cell A1 in Excel using Aspose.Cells
// Description: Creates a new workbook, inserts a rectangle shape at row 2/column 2, binds its text to cell A1 for dynamic content, refreshes the display with UpdateSelectedValue, and saves the file as RectangleLinked.xlsx.
// Keywords: Aspose.Cells C# rectangle shape | link shape to cell Excel | LinkedCell property | UpdateSelectedValue method | dynamic text in Excel shape | save workbook with shapes | Aspose.Cells .NET example
// Common Searches: how to link a rectangle shape to a cell with Aspose.Cells | Aspose.Cells C# update linked shape text | add rectangle shape to Excel worksheet using Aspose | refresh linked shape after cell change Aspose.Cells | save Excel file with shapes Aspose.Cells .NET
// Developer Intent: Insert a rectangle shape, bind its text to cell A1, refresh it, and save the workbook.
// Use Cases: Display a key metric inside a shape that updates automatically when the source cell changes. | Build a printable dashboard where shapes act as live placeholders for cell values. | Create a template that uses linked shapes to show dynamic titles or labels without manual editing.
// AI Prompts: Generate C# code to change the LinkedCell of an existing rectangle shape to B2 and refresh its value with Aspose.Cells. | Show how to add multiple shapes (rectangle, oval) each linked to different cells and ensure they update on workbook save. | Explain how to programmatically resize a linked rectangle based on the length of the text in its source cell.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, inserts a rectangle shape at row 2/column 2, binds its text to cell A1 for dynamic content, refreshes the display with UpdateSelectedValue, and saves the file as RectangleLinked.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        RectangleShape rectangle = worksheet.Shapes.AddRectangle(2, 2, 0, 0, 100, 100);

        // Link the rectangle's text to cell A1 (dynamic text)
        rectangle.LinkedCell = "$A$1";

        // Refresh the shape to display the linked cell value
        rectangle.UpdateSelectedValue();

        // Save the workbook
        workbook.Save("RectangleLinked.xlsx");
    }
}
