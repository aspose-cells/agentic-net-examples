// Title: Create a rectangle shape linked to a cell and display its value in scientific notation using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set shape text to cell value with scientific notation | how to add a rectangle shape linked to a worksheet cell using Aspose.Cells | verify shape text after saving and reloading an Excel file with Aspose.Cells | format cell value as scientific notation and assign to shape in .NET | link Excel shape to cell value programmatically with Aspose.Cells
// Tags: rectangle shape linked to worksheet cell Aspose.Cells | shape text formatted as scientific notation C# | load workbook to validate shape content Aspose.Cells | assign cell value to shape text Aspose.Cells .NET | custom number format for shape text Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, writes a numeric value to cell A1, adds a rectangle shape, sets the shape's text to the cell value formatted in scientific notation, saves the file, reloads it, and prints the shape's text to confirm the formatting.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a numeric value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(1234567.89);

            // Add a rectangle shape to the worksheet
            // Parameters: type, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = sheet.Shapes.AddShape(
                MsoDrawingType.Rectangle, // shape type
                2,    // upper left row (zero‑based)
                1,    // upper left column (zero‑based)
                0,    // top offset in pixels
                0,    // left offset in pixels
                100,  // height in pixels
                200   // width in pixels
            );

            // Format the cell value in scientific notation and assign to the shape's text
            string formatted = string.Format("{0:0.00E+00}", cell.Value);
            shape.Text = formatted;

            // ---------- Save the workbook ----------
            string filePath = "ShapeLinkedToCell.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook to verify ----------
            if (File.Exists(filePath))
            {
                Workbook loadedWorkbook = new Workbook(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                if (loadedSheet.Shapes.Count > 0)
                {
                    Shape loadedShape = loadedSheet.Shapes[0];
                    string displayedText = loadedShape.Text;
                    Console.WriteLine("Shape text (should be in scientific notation): " + displayedText);
                }
                else
                {
                    Console.WriteLine("No shapes found in the loaded workbook.");
                }
            }
            else
            {
                Console.WriteLine("Error: The file was not saved correctly.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
