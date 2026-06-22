using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the position and size of the template TextBox
        int templateRow = 2;      // upper left row index
        int templateColumn = 1;   // upper left column index
        int height = 50;          // height in pixels
        int width = 150;          // width in pixels

        // Add the template TextBox to the worksheet
        int templateIndex = sheet.TextBoxes.Add(templateRow, templateColumn, height, width);
        TextBox templateBox = sheet.TextBoxes[templateIndex];
        templateBox.Text = "Template Placeholder";
        templateBox.Font.Name = "Arial";
        templateBox.Font.Size = 12;
        templateBox.Font.IsBold = true;
        templateBox.Fill.SolidFill.Color = Color.LightGray;

        // Determine the row where the cloned TextBox will be placed (e.g., 5 rows below)
        int newRow = templateRow + 5;

        // Add a new TextBox at the new location with the same size
        int cloneIndex = sheet.TextBoxes.Add(newRow, templateColumn, height, width);
        TextBox clonedBox = sheet.TextBoxes[cloneIndex];

        // Copy visual properties from the template TextBox
        clonedBox.Font.Name = templateBox.Font.Name;
        clonedBox.Font.Size = templateBox.Font.Size;
        clonedBox.Font.IsBold = templateBox.Font.IsBold;
        clonedBox.Font.IsItalic = templateBox.Font.IsItalic;
        clonedBox.Font.Color = templateBox.Font.Color;
        clonedBox.Fill.SolidFill.Color = templateBox.Fill.SolidFill.Color;
        clonedBox.Line.Weight = templateBox.Line.Weight;
        clonedBox.Line.DashStyle = templateBox.Line.DashStyle;

        // Update the placeholder text of the cloned TextBox
        clonedBox.Text = "Cloned Placeholder";

        // Save the workbook to a file
        workbook.Save("ClonedTextBoxDemo.xlsx");
    }
}