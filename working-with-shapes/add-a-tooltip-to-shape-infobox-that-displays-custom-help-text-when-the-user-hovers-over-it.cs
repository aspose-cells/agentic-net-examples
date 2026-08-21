// Title: Add a Tooltip to a TextBox (InfoBox) Shape in Excel using Aspose.Cells for .NET
// Description: This example creates a new Workbook, inserts a TextBox shape that serves as an InfoBox, assigns custom help text via the AlternativeText property (displayed as a tooltip on hover), and saves the file as InfoBoxWithTooltip.xlsx.
// Keywords: Aspose.Cells | C# tooltip shape | Excel AlternativeText | InfoBox tooltip | Add TextBox shape | screen tip Aspose | shape hover text | .NET Excel tooltip | shape tooltip property | Aspose.Cells .NET
// Common Searches: Aspose.Cells set tooltip for shape | How to add screen tip to TextBox in Excel with C# | AlternativeText property example Aspose.Cells | Display help text on hover in Excel workbook | Add InfoBox with tooltip using Aspose.Cells
// Developer Intent: Add a custom hover tooltip with help text to a TextBox (InfoBox) shape in an Excel workbook.
// Use Cases: Provide on‑sheet help for data‑entry forms | Explain chart elements in interactive dashboards | Offer inline documentation for complex spreadsheets | Guide users through report navigation with brief notes | Show contextual tips for icons or images in Excel
// AI Prompts: Write C# code that updates the AlternativeText of an existing shape in an Aspose.Cells workbook. | Explain the difference between AlternativeText and comments for Excel shapes in Aspose.Cells and how each appears in the UI. | Show how to read, modify, and persist a shape’s tooltip after loading a workbook from disk.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// This example creates a new Workbook, inserts a TextBox shape that serves as an InfoBox, assigns custom help text via the AlternativeText property (displayed as a tooltip on hover), and saves the file as InfoBoxWithTooltip.xlsx.
class AddTooltipToInfoBox
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a textbox shape that will act as the InfoBox
        // Parameters: upper left row, upper left column, top offset, left offset, height, width
        TextBox infoBox = (TextBox)sheet.Shapes.AddTextBox(2, 1, 0, 0, 100, 200);
        infoBox.Text = "InfoBox";

        // Set the tooltip (screen tip) that appears when the user hovers over the shape
        infoBox.AlternativeText = "Custom help text displayed on hover";

        // Save the workbook
        workbook.Save("InfoBoxWithTooltip.xlsx");
    }
}
