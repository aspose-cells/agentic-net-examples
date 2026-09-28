// Title: Change a TextBox shape's linked cell from A1 to B2 and confirm the updated text using Aspose.Cells for .NET
// AI Prompts: Set the TextBox.LinkedCell property to "B2", assign a value to cell B2, call Workbook.CalculateFormula(), and output the TextBox.Text to verify the change. | Write C# code that creates a workbook, adds a TextBox linked to A1, then programmatically switches the linked cell to B2, updates the cell value, recalculates formulas, and checks the shape's displayed text.
// Common Searches: Aspose.Cells C# change TextBox linked cell from A1 to B2 | how to update shape linked cell and refresh text in Aspose.Cells .NET | verify that a linked shape reflects new cell value after recalculation in Aspose.Cells | programmatically switch linked cell of an Excel shape using Aspose.Cells | recalculate workbook after modifying shape's LinkedCell property in C#
// Tags: Aspose.Cells change TextBox linked cell | C# update shape linked cell Excel | Aspose.Cells recalculate after linked cell change | verify shape text reflects cell value Aspose.Cells | save workbook with updated shape Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Demonstrates creating a workbook, adding a TextBox shape linked to cell A1, programmatically changing its LinkedCell to B2, updating the cell value, recalculating formulas, and confirming that the shape's displayed text reflects the new linked cell before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a TextBox shape to the worksheet
            // Parameters: upperLeftRow, upperLeftColumn, height, width, rowOffset, columnOffset
            TextBox textBox = sheet.Shapes.AddTextBox(5, 5, 30, 150, 0, 0);
            textBox.Name = "LinkedTextBox";

            // Link the shape to cell A1 and set a value in A1
            textBox.LinkedCell = "A1";
            sheet.Cells["A1"].PutValue("Value in A1");

            // Recalculate to propagate the linked cell value to the shape
            workbook.CalculateFormula();

            // Verify that the shape's text reflects the value of A1
            Console.WriteLine("Linked to A1 -> Shape Text: " + textBox.Text); // Expected: "Value in A1"

            // Change the linked cell to B2 and set a value in B2
            textBox.LinkedCell = "B2";
            sheet.Cells["B2"].PutValue("Value in B2");

            // Recalculate again to update the shape
            workbook.CalculateFormula();

            // Verify that the shape's text now reflects the value of B2
            Console.WriteLine("Linked to B2 -> Shape Text: " + textBox.Text); // Expected: "Value in B2"

            // Save the workbook (optional verification step)
            string outputPath = "ShapeLinkedCellDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
