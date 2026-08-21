// Title: Aspose.Cells .NET – Set custom internal padding (margins) for a TextBox shape in Excel
// Description: Demonstrates how to add a TextBox to an Excel worksheet with Aspose.Cells, turn off automatic margin calculation, and define left, right, top, and bottom padding in points using TextBoxOptions before saving the file.
// Keywords: Aspose.Cells TextBox padding | Excel textbox internal margins .NET | TextBoxOptions margin points | disable auto margin Aspose.Cells | C# set textbox padding Excel | Aspose.Cells shape formatting | custom textbox margins Aspose
// Common Searches: Aspose.Cells set textbox padding C# | how to change internal margins of Excel textbox using Aspose | disable automatic margin calculation Aspose.Cells TextBox | set left right top bottom margins TextBoxOptions | adjust textbox padding in .xlsx with Aspose.Cells
// Developer Intent: Apply precise internal padding to a TextBox shape in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Design report templates where callout boxes need consistent spacing from their borders. | Generate invoices with note sections inside textboxes that require exact padding for readability. | Create dashboards where data labels are placed in textboxes with uniform margin settings.
// AI Prompts: Provide C# code to set different left, right, top, and bottom margins for an Aspose.Cells TextBox and keep the settings after saving. | Show how to read the current padding of an existing TextBox in an Excel file and update only the top margin using Aspose.Cells for .NET. | Explain how to toggle automatic margin calculation for a TextBox shape in Aspose.Cells and apply custom margins in points.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Demonstrates how to add a TextBox to an Excel worksheet with Aspose.Cells, turn off automatic margin calculation, and define left, right, top, and bottom padding in points using TextBoxOptions before saving the file.
class AdjustTextboxMargins
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a text box shape to the worksheet
        // Parameters: upper left row, upper left column, top offset, left offset, width, height
        TextBox textBox = sheet.Shapes.AddTextBox(1, 1, 1, 1, 200, 100);
        textBox.Text = "Text with custom internal margins (padding).";

        // Disable automatic margin calculation so custom values are applied
        textBox.TextBody.TextAlignment.IsAutoMargin = false;

        // Set internal margins (padding) in points
        textBox.TextBoxOptions.LeftMarginPt = 15;    // left padding
        textBox.TextBoxOptions.RightMarginPt = 15;   // right padding
        textBox.TextBoxOptions.TopMarginPt = 10;     // top padding
        textBox.TextBoxOptions.BottomMarginPt = 10;  // bottom padding

        // Save the workbook to a file
        workbook.Save("TextboxMarginsDemo.xlsx");
    }
}
