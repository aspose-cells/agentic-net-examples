// Title: Move an Excel TextBox to another cell by changing its anchor rows and columns with Aspose.Cells for .NET
// Common Searches: how to change the anchor row and column of a TextBox shape in Aspose.Cells C# | Aspose.Cells move textbox from B2 to D5 programmatically | set UpperLeftRow UpperLeftColumn for Excel shape using Aspose.Cells .NET | reposition a textbox shape in an .xlsx file with Aspose.Cells API | C# code to update textbox anchor coordinates in Aspose.Cells workbook
// Tags: textbox anchor reposition Aspose.Cells | update shape coordinates C# | move Excel textbox Aspose.Cells | set UpperLeftRow UpperLeftColumn Aspose.Cells | shape positioning in .xlsx with Aspose

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextboxMoveDemo
{
    // The example creates a workbook, adds a TextBox anchored to cell B2, prints its original anchor coordinates, then modifies UpperLeftRow, UpperLeftColumn, LowerRightRow, and LowerRightColumn to move the TextBox to cell D5, updates its text, displays the new coordinates, and saves the file as TextboxMoved.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "TextboxMoved.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Add a textbox anchored to cell B2 (row 1, column 1)
                //    Returns a TextBox object (derived from Shape)
                TextBox textBox = sheet.Shapes.AddTextBox(1, 1, 0, 0, 150, 50);

                // Optional visual styling
                textBox.Text = "Original Position (B2)";
                textBox.Font.Color = Color.White;
                textBox.Font.Size = 12;
                // Set fill pattern (ForeColor may not be supported in all versions)
                textBox.Fill.Pattern = FillPattern.Solid;
                // If ForeColor is available, uncomment the following line:
                // textBox.Fill.ForeColor = Color.DarkBlue;

                // 4. Display the original anchor coordinates
                Console.WriteLine("Original Anchor:");
                Console.WriteLine($" UpperLeftRow    = {textBox.UpperLeftRow}");
                Console.WriteLine($" UpperLeftColumn = {textBox.UpperLeftColumn}");
                Console.WriteLine($" LowerRightRow   = {textBox.LowerRightRow}");
                Console.WriteLine($" LowerRightColumn= {textBox.LowerRightColumn}");

                // 5. Move the textbox to a new location (cell D5)
                textBox.UpperLeftRow = 4;      // D5 row (zero‑based)
                textBox.UpperLeftColumn = 3;   // D5 column
                textBox.LowerRightRow = 6;     // two rows below
                textBox.LowerRightColumn = 5;  // two columns to the right

                // Update displayed text
                textBox.Text = "Moved Position (D5)";

                // 6. Display the new anchor coordinates
                Console.WriteLine("\nAfter Move Anchor:");
                Console.WriteLine($" UpperLeftRow    = {textBox.UpperLeftRow}");
                Console.WriteLine($" UpperLeftColumn = {textBox.UpperLeftColumn}");
                Console.WriteLine($" LowerRightRow   = {textBox.LowerRightRow}");
                Console.WriteLine($" LowerRightColumn= {textBox.LowerRightColumn}");

                // 7. Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while moving the textbox:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
