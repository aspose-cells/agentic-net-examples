// Title: Add a TextBox Shape to a Worksheet at Specific Cell Coordinates with Aspose.Cells for .NET
// Description: Shows how to create a workbook, locate a cell using zero‑based row and column indices, insert a TextBox shape of defined pixel dimensions, apply text and font styling, and save the Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# add textbox | textbox shape | cell coordinates | row column indices | shape positioning | font styling | .NET Excel automation | TextBoxCollection
// Common Searches: Aspose.Cells add textbox to worksheet | position textbox by row and column | set textbox size in pixels Aspose.Cells | change textbox font color C# | save workbook after adding shape Aspose.Cells | Aspose.Cells TextBoxCollection example
// Developer Intent: Insert a TextBox shape at a specified row and column and customize its appearance in an Excel workbook using Aspose.Cells.
// Use Cases: Add a labeled instruction box at a header row for consistent report layouts. | Create a dynamic title box that stays aligned with a specific cell range across different worksheets. | Highlight key metrics with a callout box positioned next to the target data cell.
// AI Prompts: Write C# code with Aspose.Cells to add a textbox at row 5, column 3, using red italic font and custom text. | Explain how to convert column width and row height to pixel dimensions for accurately sizing a textbox in Aspose.Cells. | Provide steps to move and resize an existing textbox programmatically after it has been added to a worksheet.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to create a workbook, locate a cell using zero‑based row and column indices, insert a TextBox shape of defined pixel dimensions, apply text and font styling, and save the Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Specify the cell coordinates (zero‑based) where the textbox will be placed
        int topRow = 2;      // Row index (e.g., third row)
        int leftColumn = 1;  // Column index (e.g., second column)

        // Define the size of the textbox in pixels
        int height = 80;
        int width = 200;

        // Add a textbox to the worksheet using the TextBoxCollection
        int textboxIndex = worksheet.TextBoxes.Add(topRow, leftColumn, height, width);
        TextBox textbox = worksheet.TextBoxes[textboxIndex];

        // Set some properties of the textbox
        textbox.Text = "Hello from Aspose.Cells!";
        textbox.Font.Size = 12;
        textbox.Font.IsBold = true;
        textbox.Font.Color = Color.Blue;

        // Save the workbook to a file
        workbook.Save("TextboxDemo.xlsx");
    }
}
