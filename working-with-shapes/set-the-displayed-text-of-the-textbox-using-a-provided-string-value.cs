// Title: How to set the displayed text of a TextBox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add a TextBox shape at a specific cell location, and assign a custom string to its Text property with Aspose.Cells in C#. | Update the Text property of an existing TextBox shape, apply bold formatting to its font, and save the workbook using Aspose.Cells. | Write C# code that inserts a TextBox into the first worksheet, sets the displayed text from a variable, and exports the file as an .xlsx document.
// Common Searches: Aspose.Cells C# set TextBox shape text programmatically | How to change the displayed string of a TextBox in an Excel file using Aspose.Cells | C# example for adding a TextBox with bold text to a worksheet with Aspose.Cells | Saving a workbook after updating TextBox content with Aspose.Cells .NET
// Tags: Aspose.Cells set TextBox Text property C# | add TextBox shape to worksheet Aspose.Cells | format TextBox font bold Aspose.Cells | save workbook as XLSX Aspose.Cells | initialize TextBox position rows columns Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, inserting a TextBox shape at a defined position, assigning a custom string to its Text property, optionally applying bold formatting, and saving the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define textbox position and size
        int upperLeftRow = 2;      // Row index (0‑based)
        int upperLeftColumn = 1;   // Column index (0‑based)
        int top = 5;               // Top offset in pixels
        int left = 5;              // Left offset in pixels
        int height = 100;          // Height in pixels
        int width = 200;           // Width in pixels

        // Add a textbox shape to the worksheet
        TextBox textbox = sheet.Shapes.AddTextBox(
            upperLeftRow, upperLeftColumn, top, left, height, width);

        // Set the displayed text of the textbox
        string displayText = "Hello, Aspose.Cells!";
        textbox.Text = displayText;

        // (Optional) Apply simple formatting
        textbox.Font.IsBold = true;
        textbox.Font.Size = 12;

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
