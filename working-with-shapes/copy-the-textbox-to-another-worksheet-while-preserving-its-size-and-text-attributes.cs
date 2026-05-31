using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class CopyTextBoxExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Source worksheet
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "Source";

        // Add a TextBox to the source worksheet
        TextBox textBox = sourceSheet.Shapes.AddTextBox(2, 0, 2, 0, 200, 100);
        textBox.Text = "Hello Aspose.Cells!";
        textBox.Font.Name = "Arial";
        textBox.Font.Size = 12;
        textBox.Font.IsBold = true;

        // Destination worksheet
        Worksheet destSheet = workbook.Worksheets.Add("Destination");

        // Copy the TextBox to the destination worksheet preserving size and text attributes
        Shape sourceShape = sourceSheet.Shapes[0]; // the TextBox we just added
        destSheet.Shapes.AddCopy(
            sourceShape,
            sourceShape.UpperLeftRow,   // top row index
            sourceShape.Top,            // vertical offset in pixels
            sourceShape.UpperLeftColumn,// left column index
            sourceShape.Left);          // horizontal offset in pixels

        // Save the workbook
        workbook.Save("CopyTextBoxResult.xlsx");
    }
}