// Title: Aspose.Cells C# – Relocate a TextBox by Setting UpperLeftRow & UpperLeftColumn
// Description: Shows how to create a workbook, add a TextBox at cell B3 (row 2, column 2), move it to cell C6 (row 5, column 3) by updating the UpperLeftRow and UpperLeftColumn properties, optionally resize it, and save the result as MovedTextBox.xlsx.
// Keywords: Aspose.Cells move textbox C# | Aspose.Cells TextBox anchor | UpperLeftRow UpperLeftColumn | change shape position Aspose.Cells | C# reposition TextBox | Aspose.Cells worksheet shapes | adjust TextBox size Aspose.Cells
// Common Searches: Aspose.Cells move textbox to another cell | How to change TextBox anchor row and column in Aspose.Cells .NET | C# set UpperLeftRow UpperLeftColumn for shape | Resize TextBox after moving Aspose.Cells | Move multiple TextBoxes programmatically Aspose.Cells
// Developer Intent: Reposition an existing TextBox to a different cell by modifying its UpperLeftRow and UpperLeftColumn anchor values.
// Use Cases: Keep a label TextBox aligned with a data column after rows are inserted or deleted. | Place a TextBox next to a dynamically generated table header by anchoring it to the header cell. | Customize a worksheet template by moving several TextBoxes to user‑specified cells during report generation.
// AI Prompts: Generate C# code with Aspose.Cells that moves a TextBox from row 2, column 2 to row 10, column 5 while preserving its original height and width. | Explain how UpperLeftRow and UpperLeftColumn define a TextBox's anchor point and how to calculate them for merged cells in Aspose.Cells. | Provide a C# loop that iterates over a dictionary of cell coordinates and repositions each TextBox in a worksheet accordingly.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, add a TextBox at cell B3 (row 2, column 2), move it to cell C6 (row 5, column 3) by updating the UpperLeftRow and UpperLeftColumn properties, optionally resize it, and save the result as MovedTextBox.xlsx.
class MoveTextBoxDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a textbox at an initial location (row 2, column 2)
        int textboxIndex = worksheet.TextBoxes.Add(2, 2, 100, 200);
        TextBox textbox = worksheet.TextBoxes[textboxIndex];
        textbox.Text = "Original Position";

        // Move the textbox to a new location (row 5, column 3) by updating its anchor coordinates
        textbox.UpperLeftRow = 5;
        textbox.UpperLeftColumn = 3;

        // Optionally adjust size after moving
        textbox.Height = 120;
        textbox.Width = 250;

        // Save the workbook
        workbook.Save("MovedTextBox.xlsx");
    }
}
