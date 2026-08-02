// Title: C# – Add a TextBox to Sheet1 at B2 and set Width to 200 pt with Aspose.Cells
// Description: Demonstrates how to create a workbook, access Sheet1, insert a TextBox anchored to cell B2, change its width to 200 points using the WidthPt property, optionally set its text, and save the file as Output.xlsx.
// Keywords: Aspose.Cells C# textbox | add textbox to worksheet | set textbox width points | position shape at B2 | Aspose.Cells shape sizing | Excel automation Aspose.Cells
// Common Searches: Aspose.Cells add textbox at specific cell | C# set textbox width in points Aspose.Cells | How to position a shape on B2 using Aspose.Cells | Change textbox dimensions Aspose.Cells .NET | Aspose.Cells TextBox WidthPt example
// Developer Intent: Insert a TextBox on Sheet1 at cell B2 and define its width as 200 pt.
// Use Cases: Add instructional notes to a generated report at a fixed cell location. | Create placeholder fields for later data population in automated Excel files. | Design form‑like layouts by aligning multiple TextBoxes to specific worksheet cells.
// AI Prompts: Write C# code with Aspose.Cells to place a TextBox at cell C3, height 150 pt, width 250 pt. | Show how to align a column of TextBoxes vertically on a worksheet using Aspose.Cells. | Explain converting between points and pixels for shape dimensions in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, access Sheet1, insert a TextBox anchored to cell B2, change its width to 200 points using the WidthPt property, optionally set its text, and save the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet (Sheet1)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox positioned at cell B2 (row index 1, column index 1)
        // Parameters: topRow, leftColumn, height (pixels), width (pixels)
        int textboxIndex = sheet.TextBoxes.Add(1, 1, 100, 100);
        TextBox textbox = sheet.TextBoxes[textboxIndex];

        // Set the width of the textbox to 200 points
        textbox.WidthPt = 200;

        // Optional: set some sample text
        textbox.Text = "Sample TextBox";

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
