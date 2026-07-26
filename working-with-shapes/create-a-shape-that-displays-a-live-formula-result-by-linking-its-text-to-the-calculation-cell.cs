// Title: C# – Link a TextBox Shape to a Cell for Live Formula Results with Aspose.Cells
// Description: Creates a workbook, fills A1:A5, adds a SUM formula in B2, inserts a TextBox shape, links it to B2 using SetLinkedCell so the shape updates automatically, and saves the file as LinkedShapeDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | TextBox shape | linked cell | live formula result | Excel automation | SetLinkedCell API | dynamic shape text | sample code | GitHub example
// Common Searches: Aspose.Cells link textbox to cell C# | display live formula value in shape Aspose.Cells | how to bind a shape to a cell in Excel using .NET | update textbox text automatically when cell changes Aspose.Cells | sample code for SetLinkedCell method
// Developer Intent: Show how to bind a TextBox shape to a worksheet cell so the shape reflects the current formula result.
// Use Cases: Dashboard total displayed in a shape that updates with data changes. | Report header that shows a calculated date or metric without manual edits. | Interactive template where shape captions stay in sync with input values.
// AI Prompts: Generate C# code that adds a rectangle shape linked to cell C5 to show a VLOOKUP result using Aspose.Cells. | Explain how to link multiple shapes to different cells and format each shape after linking. | Provide a step‑by‑step guide to change font style and alignment of a linked TextBox shape in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, fills A1:A5, adds a SUM formula in B2, inserts a TextBox shape, links it to B2 using SetLinkedCell so the shape updates automatically, and saves the file as LinkedShapeDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some cells with numbers
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["A4"].PutValue(40);
        sheet.Cells["A5"].PutValue(50);

        // Set a formula in B2 that sums A1:A5
        sheet.Cells["B2"].Formula = "SUM(A1:A5)";

        // Add a textbox shape that will display the live result of B2
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        TextBox txtBox = sheet.Shapes.AddTextBox(2, 2, 0, 0, 150, 30);

        // Link the textbox to cell B2 so its displayed text updates automatically
        txtBox.SetLinkedCell("$B$2", false, true);

        // Optional: clear any placeholder text
        txtBox.Text = string.Empty;

        // Save the workbook
        workbook.Save("LinkedShapeDemo.xlsx");
    }
}
