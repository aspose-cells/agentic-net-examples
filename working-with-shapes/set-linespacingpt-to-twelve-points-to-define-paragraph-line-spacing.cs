// Title: Set 12‑point paragraph line spacing in a text box shape with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a text box, select a specific paragraph, set its LineSpaceSizeType to Points and LineSpace to 12, and save the Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells line spacing | C# set paragraph line spacing | text box line spacing points | LineSpaceSizeType Points | Aspose.Cells shape text formatting | Excel text box paragraph spacing | Aspose.Cells .NET paragraph formatting
// Common Searches: Aspose.Cells set paragraph line spacing | C# line spacing points in Excel text box | How to change line spacing of shape text in Aspose.Cells | Set line spacing to 12 points Aspose.Cells .NET | Adjust text box paragraph spacing programmatically
// Developer Intent: Configure the line spacing of a selected paragraph inside a worksheet text box to exactly 12 points with Aspose.Cells for .NET.
// Use Cases: Generate a report workbook where a text box paragraph follows a 12‑point line‑spacing rule for consistent appearance. | Apply corporate style guidelines by programmatically enforcing fixed point line spacing on shape text across multiple worksheets. | Update existing Excel files to standardize paragraph spacing inside text boxes without manual editing.
// AI Prompts: Write C# code using Aspose.Cells to set the line spacing of the third paragraph in a shape to 10 points. | Provide an example that changes the line‑spacing type to Points and sets it to 14 for every paragraph in all text boxes on a worksheet. | Explain how to read, modify, and save line‑spacing values for each paragraph in a shape's TextBody with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Shows how to create a workbook, add a text box, select a specific paragraph, set its LineSpaceSizeType to Points and LineSpace to 12, and save the Excel file using Aspose.Cells for .NET.
class SetParagraphLineSpacing
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        Shape textBox = sheet.Shapes.AddTextBox(0, 0, 0, 0, 400, 200);
        textBox.Text = "First line\nSecond line";

        // Access the second paragraph (index 1) of the text box
        TextParagraph paragraph = textBox.TextBody.TextParagraphs[1];

        // Define line spacing in points and set it to 12 points
        paragraph.LineSpaceSizeType = LineSpaceSizeType.Points;
        paragraph.LineSpace = 12;

        // Save the workbook to a file
        workbook.Save("LineSpacingDemo.xlsx");
    }
}
