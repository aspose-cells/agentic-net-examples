// Title: Create and configure a TextBox shape with initial text on the first worksheet using Aspose.Cells for .NET
// Common Searches: c# add textbox shape to Aspose.Cells worksheet and set its text | how to set textbox font and placement in Aspose.Cells .NET | Aspose.Cells add TextBox at specific row and column example | move and size textbox with cells using Aspose.Cells C# | save workbook after inserting a textbox shape Aspose.Cells | Aspose.Cells TextBox placement MoveAndSize property usage
// Tags: Aspose.Cells TextBox shape insertion | Aspose.Cells TextBox initial text setting | Aspose.Cells TextBox font styling | Aspose.Cells TextBox placement MoveAndSize | Aspose.Cells workbook save with shapes

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TextBoxDemoApp
{
    // The example creates a new Workbook, accesses the first Worksheet, inserts a TextBox at row 2 column 1 with defined pixel dimensions, assigns the text "Hello, Aspose.Cells!", customizes the font (name, size, bold, color), sets the placement to MoveAndSize so the shape follows cell changes, and saves the file as TextBoxDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Get the first worksheet.
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a TextBox to the worksheet.
                // Parameters: firstRow, firstColumn, width (pixels), height (pixels)
                // Here we place it starting at row 2 (index 1), column 1 (index 0).
                int textboxIndex = worksheet.TextBoxes.Add(1, 0, 200, 80);

                // Retrieve the added TextBox.
                TextBox textbox = worksheet.TextBoxes[textboxIndex];

                // Set the initial text of the TextBox.
                textbox.Text = "Hello, Aspose.Cells!";

                // Optional: adjust some visual properties.
                textbox.Font.Name = "Calibri";
                textbox.Font.Size = 12;
                textbox.Font.IsBold = true;
                textbox.Font.Color = Color.DarkBlue;

                // Set the placement so the textbox moves and resizes with cells.
                textbox.Placement = PlacementType.MoveAndSize;

                // Save the workbook to a file.
                string outputPath = "TextBoxDemo.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
