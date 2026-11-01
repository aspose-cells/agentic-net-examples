// Title: Duplicate a textbox shape, link the copy to another cell, and compare their text using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to copy an existing textbox shape to cell A6, set its Placement to MoveAndSize, and verify whether the copied shape's Text property equals the original. | Create a .NET example that adds a textbox linked to cell A1, duplicates it with AddCopy to a different cell, and prints a message indicating if the two shapes contain the same text.
// Common Searches: Aspose.Cells C# duplicate shape and link to another cell example | How to use AddCopy to copy a textbox shape to a different location in a worksheet | Compare text of two shapes in an Aspose.Cells workbook using C# | Set shape placement to MoveAndSize after copying with Aspose.Cells | Save workbook after duplicating shapes with Aspose.Cells .NET
// Tags: Aspose.Cells AddCopy shape duplication | textbox shape linked to cell placement | compare shape Text property C# | MoveAndSize placement for shapes | duplicate shape in worksheet .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDuplicationDemoApp
{
    // The sample creates a workbook, adds a textbox shape linked to cell A1, duplicates the shape to cell A6 using the AddCopy method, sets both shapes to MoveAndSize placement, compares their Text values, outputs the comparison result, and saves the file as ShapeDuplicationResult.xlsx.
    class ShapeDuplicationDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a textbox shape linked to cell A1 (row 0, column 0)
                Shape originalShape = sheet.Shapes.AddTextBox(0, 0, 0, 0, 150, 40);
                originalShape.Text = "Original Shape";
                // Ensure the shape moves and resizes with its linked cell
                originalShape.Placement = PlacementType.MoveAndSize;

                // Duplicate the shape within the same worksheet.
                // Provide the destination cell range (A6:A6) for the copy.
                Shape duplicatedShape = sheet.Shapes.AddCopy(originalShape, 5, 0, 5, 0);
                // Position the duplicate at cell A6 (row 5, column 0)
                duplicatedShape.UpperLeftRow = 5;
                duplicatedShape.UpperLeftColumn = 0;
                duplicatedShape.Placement = PlacementType.MoveAndSize;

                // Compare the text of the two shapes
                bool contentsAreEqual = originalShape.Text == duplicatedShape.Text;

                // Output the comparison result
                Console.WriteLine($"Original shape text: \"{originalShape.Text}\"");
                Console.WriteLine($"Duplicated shape text: \"{duplicatedShape.Text}\"");
                Console.WriteLine($"Contents are {(contentsAreEqual ? "identical" : "different")}.");

                // Save the workbook
                string outputPath = "ShapeDuplicationResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
