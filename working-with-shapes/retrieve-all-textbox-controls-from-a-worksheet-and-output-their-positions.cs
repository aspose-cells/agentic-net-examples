using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class RetrieveTextBoxes
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access the collection of TextBox objects on the worksheet
        TextBoxCollection textBoxes = worksheet.TextBoxes;

        // Iterate through all TextBoxes and output their positions
        for (int i = 0; i < textBoxes.Count; i++)
        {
            TextBox tb = textBoxes[i];

            Console.WriteLine($"TextBox {i}:");
            Console.WriteLine($"  UpperLeftRow    : {tb.UpperLeftRow}");
            Console.WriteLine($"  UpperLeftColumn : {tb.UpperLeftColumn}");
            Console.WriteLine($"  LowerRightRow   : {tb.LowerRightRow}");
            Console.WriteLine($"  LowerRightColumn: {tb.LowerRightColumn}");
            Console.WriteLine($"  Top (pixels)    : {tb.Top}");
            Console.WriteLine($"  Left (pixels)   : {tb.Left}");
            Console.WriteLine($"  Height (pixels) : {tb.Height}");
            Console.WriteLine($"  Width (pixels)  : {tb.Width}");
            Console.WriteLine();
        }

        // Save the workbook (no modifications made, just to follow lifecycle rules)
        workbook.Save("output.xlsx");
    }
}