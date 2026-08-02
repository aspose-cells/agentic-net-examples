// Title: Clone a TextBox, Change Its Text, and Insert into a New Row with Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds a formatted template TextBox, inserts a new worksheet row, clones the TextBox into that row, updates the cloned placeholder text, and saves the file as an XLSX document using Aspose.Cells.
// Keywords: Aspose.Cells clone textbox C# | Aspose.Cells copy shape | Aspose.Cells insert row with shape | Aspose.Cells update textbox text | Aspose.Cells .NET TextBox example | Aspose.Cells shape formatting | Aspose.Cells worksheet shapes
// Common Searches: how to duplicate a TextBox in Aspose.Cells | clone textbox and change text Aspose.Cells .NET | insert a shape into a new row using Aspose.Cells | copy textbox formatting Aspose.Cells C# | Aspose.Cells add TextBox to worksheet programmatically
// Developer Intent: The developer needs to programmatically copy an existing TextBox, modify its placeholder text, and place the copy in a newly inserted row of an Excel worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a report where each data row includes a styled comment box derived from a template. | Automate invoice creation with identical label TextBoxes across rows, updating only the displayed text per line item. | Build a dynamic form where a template TextBox is cloned for every newly added record while preserving font and fill settings.
// AI Prompts: Show me C# code to clone a TextBox shape, preserve its formatting, insert a new row, and set new placeholder text with Aspose.Cells. | Provide a reusable method that takes a source TextBox and a target row index, inserts the row, clones the TextBox into that row, and updates its content. | Explain how to copy font, color, and size properties of a TextBox when moving it to another location in an Aspose.Cells worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds a formatted template TextBox, inserts a new worksheet row, clones the TextBox into that row, updates the cloned placeholder text, and saves the file as an XLSX document using Aspose.Cells.
class CloneTextBoxExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Create a template TextBox (this would normally
        //    already exist in the sheet; we create it here
        //    for demonstration purposes)
        // -------------------------------------------------
        // Parameters: topRow, leftColumn, height (pixels), width (pixels)
        int templateIndex = worksheet.TextBoxes.Add(2, 2, 50, 150);
        TextBox templateBox = worksheet.TextBoxes[templateIndex];
        templateBox.Text = "Template Placeholder";
        // Example of setting additional formatting (optional)
        templateBox.Font.Name = "Arial";
        templateBox.Font.Size = 12;
        templateBox.Font.IsBold = true;

        // -------------------------------------------------
        // 2. Insert a new row where the cloned TextBox will be placed
        // -------------------------------------------------
        // Insert one row after row index 5 (zero‑based, so row 6 in Excel)
        int targetRow = 6;
        worksheet.Cells.InsertRows(targetRow, 1);

        // -------------------------------------------------
        // 3. Clone the template TextBox
        // -------------------------------------------------
        // Add a new TextBox at the same column as the template,
        // using the same size as the template.
        int clonedIndex = worksheet.TextBoxes.Add(targetRow, 2, templateBox.Height, templateBox.Width);
        TextBox clonedBox = worksheet.TextBoxes[clonedIndex];

        // Copy desired properties from the template.
        // Here we copy the font settings and fill color as an example.
        clonedBox.Font.Name = templateBox.Font.Name;
        clonedBox.Font.Size = templateBox.Font.Size;
        clonedBox.Font.IsBold = templateBox.Font.IsBold;
        clonedBox.Font.Color = templateBox.Font.Color;

        // Update the placeholder text in the cloned TextBox
        clonedBox.Text = "Cloned TextBox Content";

        // -------------------------------------------------
        // 4. Save the workbook
        // -------------------------------------------------
        workbook.Save("ClonedTextBox.xlsx");
    }
}
