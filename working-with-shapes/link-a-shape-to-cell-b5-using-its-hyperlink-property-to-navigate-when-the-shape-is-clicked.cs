// Title: Aspose.Cells C# – Link a Shape to Cell B5 Using the Hyperlink Property
// Description: This example creates a workbook, adds a rectangle shape to the first worksheet, assigns a Hyperlink with the address "Sheet1!B5" and a ScreenTip to the shape, and saves the file. Clicking the shape navigates directly to cell B5.
// Keywords: Aspose.Cells | C# | shape hyperlink | cell navigation | rectangle shape | worksheet API | hyperlink address | B5 | programmatic linking | .NET Excel
// Common Searches: Aspose.Cells set shape hyperlink to a cell | C# link rectangle shape to B5 in Excel | How to make a shape navigate to a specific cell with Aspose.Cells | Add clickable shape that jumps to a worksheet cell
// Developer Intent: Programmatically assign a hyperlink to a shape so that a click jumps to cell B5 on the same sheet.
// Use Cases: Create a dashboard button that opens a summary cell instantly. | Add navigation markers in a multi‑section report for quick jumps. | Embed interactive diagram elements that point to underlying data cells.
// AI Prompts: Generate C# code with Aspose.Cells to add a shape and link it to cell C10, including a custom ScreenTip. | Explain how to set a shape's hyperlink to a different worksheet or an external URL using Aspose.Cells. | Show how to retrieve an existing shape's Hyperlink object, modify its address, and update the ScreenTip.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds a rectangle shape to the first worksheet, assigns a Hyperlink with the address "Sheet1!B5" and a ScreenTip to the shape, and saves the file. Clicking the shape navigates directly to cell B5.
    public class ShapeLinkToCellDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, top, left, height, width
                Shape shape = sheet.Shapes.AddRectangle(1, 1, 0, 0, 100, 150);

                // Set hyperlink to cell B5 on Sheet1
                Hyperlink hyperlink = shape.Hyperlink;
                hyperlink.Address = "Sheet1!B5";
                hyperlink.ScreenTip = "Click to go to B5";

                // Save the workbook
                workbook.Save("ShapeLinkToCell.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ShapeLinkToCellDemo.Run();
        }
    }
}
