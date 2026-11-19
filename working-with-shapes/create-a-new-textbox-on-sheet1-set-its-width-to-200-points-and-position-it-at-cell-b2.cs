// Title: Create a 200‑point wide TextBox shape anchored to cell B2 on Sheet1 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add TextBox shape at cell B2 with specific width | set textbox width to 200 points in Aspose.Cells workbook | how to position a TextBox on a worksheet cell using Aspose.Cells .NET | Aspose.Cells shape dimensions example C# | anchor TextBox to a cell in Aspose.Cells and define size
// Tags: Aspose.Cells add TextBox shape | TextBox shape width points Aspose.Cells | anchor shape to worksheet cell Aspose.Cells | C# Aspose.Cells shape positioning | Aspose.Cells shape dimensions setting | Aspose.Cells workbook save with shapes

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;

// The example creates a new Workbook, accesses Sheet1, adds a TextBox shape anchored to cell B2 with a height of 100 points and a width of 200 points, sets sample text, and saves the file as output.xlsx, demonstrating shape creation, sizing, and positioning with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (Sheet1)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Add a TextBox shape positioned at cell B2 (row index 1, column index 1)
            // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
            // Top and left offsets are set to 0, height is arbitrary (e.g., 100 points)
            TextBox textbox = sheet.Shapes.AddTextBox(1, 1, 0, 0, 100, 200);

            // Ensure the width is exactly 200 points
            textbox.Width = 200;

            // Optional: set some text inside the TextBox
            textbox.Text = "Sample TextBox";

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
