using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add two overlapping text boxes
        int tbIndex1 = sheet.TextBoxes.Add(5, 5, 150, 80);
        TextBox textBox1 = sheet.TextBoxes[tbIndex1];
        textBox1.Text = "Background Box";

        int tbIndex2 = sheet.TextBoxes.Add(30, 30, 150, 80);
        TextBox textBox2 = sheet.TextBoxes[tbIndex2];
        textBox2.Text = "Foreground Box";

        // Bring the second text box to the front of the Z‑order
        textBox2.ToFrontOrBack(1); // Positive value moves the shape forward

        // Save the workbook
        workbook.Save("ZOrderTextBoxDemo.xlsx");
    }
}