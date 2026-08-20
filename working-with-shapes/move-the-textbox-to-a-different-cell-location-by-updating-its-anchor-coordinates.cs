// Title: Move a TextBox to a New Cell in Aspose.Cells (C#)
// Description: Creates a workbook, adds a TextBox at row 2 col 1, sets its text, then repositions it to row 5 col 3 by updating UpperLeftRow and UpperLeftColumn before saving the file.
// Keywords: Aspose.Cells | C# | move textbox | textbox position | UpperLeftRow | UpperLeftColumn | shape anchor | Excel worksheet
// Common Searches: Aspose.Cells change textbox location C# | How to set UpperLeftRow UpperLeftColumn for a textbox | Relocate textbox to another cell in Aspose.Cells | Update textbox anchor coordinates in .NET | Move shape to different row and column Aspose.Cells
// Developer Intent: Reposition an existing TextBox by modifying its UpperLeftRow and UpperLeftColumn properties.
// Use Cases: Place a label next to a dynamically generated table in a report. | Adjust textbox placement after inserting rows or columns. | Standardize textbox alignment across multiple dashboard sheets.
// AI Prompts: Generate C# code that moves a textbox to row 10, column 2 and resizes it using Aspose.Cells. | Show how to shift all textboxes down by two rows in a worksheet with Aspose.Cells. | Explain the role of UpperLeftRow and UpperLeftColumn in determining textbox placement in an Excel file created with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a workbook, adds a TextBox at row 2 col 1, sets its text, then repositions it to row 5 col 3 by updating UpperLeftRow and UpperLeftColumn before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a textbox at initial position (row 2, column 1) with height 100px and width 200px
                int textboxIndex = worksheet.TextBoxes.Add(2, 1, 100, 200);

                // Retrieve the added textbox
                TextBox textbox = worksheet.TextBoxes[textboxIndex];

                // Set sample text
                textbox.Text = "Sample TextBox";

                // Move the textbox to a new location (upper‑left corner at row 5, column 3)
                textbox.UpperLeftRow = 5;
                textbox.UpperLeftColumn = 3;

                // Define output file path
                string outputPath = "MovedTextbox.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
