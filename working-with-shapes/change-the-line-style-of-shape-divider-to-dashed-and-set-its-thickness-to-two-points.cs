using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the shape named "Divider"
        Shape divider = worksheet.Shapes["Divider"]; // or worksheet.Shapes.GetByName("Divider");

        if (divider != null)
        {
            // Access the line formatting of the shape
            LineFormat lineFormat = divider.Line;

            // Set the dash style to dashed
            lineFormat.DashStyle = MsoLineDashStyle.Dash;

            // Set the line thickness to 2 points
            lineFormat.Weight = 2.0;
        }
        else
        {
            Console.WriteLine("Shape 'Divider' not found.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}