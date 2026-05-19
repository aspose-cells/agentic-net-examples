using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtColorChange
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the collection of shapes on the current worksheet
                ShapeCollection shapes = sheet.Shapes;

                // Loop through all shapes
                foreach (Shape shape in shapes)
                {
                    // Identify WordArt shapes using the IsWordArt property
                    if (shape.IsWordArt)
                    {
                        // Change the text color to dark blue
                        shape.Font.Color = Color.FromArgb(0, 0, 139); // DarkBlue
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("Output.xlsx");
        }
    }
}