using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Get the first worksheet
        var sheet = workbook.Worksheets[0];

        // Populate sample data in column B (B1:B5)
        sheet.Cells["B1"].PutValue("Apple");
        sheet.Cells["B2"].PutValue("Banana");
        sheet.Cells["B3"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue("Date");
        sheet.Cells["B5"].PutValue("Elderberry");

        // Set a TEXTJOIN formula in cell A1 to concatenate the values from B1:B5
        sheet.Cells["A1"].Formula = "TEXTJOIN(\", \", TRUE, B1:B5)";

        // Add a TextBox shape to the worksheet
        // Parameters: upper left row, upper left column, row offset, column offset, width, height
        var shape = sheet.Shapes.AddTextBox(2, 0, 0, 0, 200, 50);

        // Link the shape to cell A1 so it displays the result of the TEXTJOIN formula
        shape.SetLinkedCell("A1", true, true);

        // Clear any static text; the shape will show the linked cell's value
        shape.Text = "";

        // Save the workbook
        workbook.Save("LinkedShape.xlsx");
    }
}