using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class DeleteTextboxDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox to the worksheet
        worksheet.TextBoxes.Add(5, 5, 200, 50);
        worksheet.TextBoxes[0].Text = "Sample TextBox";

        Console.WriteLine("Before removal - TextBox count: " + worksheet.TextBoxes.Count);

        // Remove the textbox using RemoveAt (index of the textbox to delete)
        int indexToRemove = 0; // remove the first textbox
        worksheet.TextBoxes.RemoveAt(indexToRemove);

        Console.WriteLine("After removal - TextBox count: " + worksheet.TextBoxes.Count);

        // Save the workbook (optional)
        workbook.Save("DeleteTextboxDemo.xlsx");

        // Release resources
        workbook.Dispose();
    }
}