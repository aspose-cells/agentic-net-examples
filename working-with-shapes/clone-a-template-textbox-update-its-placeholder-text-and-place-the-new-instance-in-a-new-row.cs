// Title: Clone a TextBox Template, Update Placeholder, and Insert into a New Row with Aspose.Cells for .NET
// Description: Demonstrates how to add a template TextBox to a worksheet, clone it to a newly inserted row, copy its visual properties, replace the placeholder text, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | clone TextBox Aspose.Cells | copy TextBox formatting | insert row Excel Aspose | placeholder text TextBox | shape duplication Aspose.Cells | Excel workbook TextBox styling | Aspose.Cells Drawing API | programmatic TextBox creation | Excel shape copy C#
// Common Searches: how to duplicate a TextBox in Aspose.Cells | copy TextBox style to another row .NET | replace placeholder text in Aspose.Cells TextBox | insert a new row and add a shape with Aspose.Cells | clone shape properties Aspose.Cells C# | Aspose.Cells TextBox example
// Developer Intent: Create a reusable TextBox template, clone it to a new worksheet row, preserve its formatting, and set custom text.
// Use Cases: Generate a series of employee name labels, each in its own styled TextBox on separate rows. | Build a fillable form where a placeholder TextBox is programmatically replaced with actual data for each record. | Maintain consistent layout in reports by duplicating a formatted TextBox across multiple rows.
// AI Prompts: Show C# code that clones a TextBox shape, copies all font, fill, and line settings, and places it in a newly inserted row using Aspose.Cells. | Provide an example that iterates over a data table and creates a cloned TextBox for each row, updating the placeholder with the row's value. | Explain how to copy visual properties from one Aspose.Cells TextBox to another, including background color and border style.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

// Demonstrates how to add a template TextBox to a worksheet, clone it to a newly inserted row, copy its visual properties, replace the placeholder text, and save the workbook using Aspose.Cells for .NET.
class CloneTextBoxExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Add a template TextBox (acts as the source)
        // -------------------------------------------------
        // Parameters: topRow, leftColumn, height (pixels), width (pixels)
        int templateTopRow = 2;      // row index where template starts
        int templateLeftColumn = 1; // column index where template starts
        int templateHeight = 60;    // height in pixels
        int templateWidth = 200;    // width in pixels

        int templateIndex = worksheet.TextBoxes.Add(templateTopRow, templateLeftColumn, templateHeight, templateWidth);
        TextBox templateBox = worksheet.TextBoxes[templateIndex];

        // Set placeholder text in the template
        templateBox.Text = "{{Placeholder}}";
        templateBox.Font.Name = "Calibri";
        templateBox.Font.Size = 12;
        templateBox.Font.IsBold = true;
        templateBox.Font.Color = Color.DarkBlue;

        // -------------------------------------------------
        // 2. Clone the template TextBox to a new row
        // -------------------------------------------------
        // Determine the row where the new TextBox will be placed
        // Insert a new row below the template row to keep layout tidy
        int newRowIndex = templateTopRow + 5; // arbitrary offset for demonstration
        worksheet.Cells.InsertRow(newRowIndex);

        // Add a new TextBox using the same size and column as the template
        int newIndex = worksheet.TextBoxes.Add(newRowIndex, templateLeftColumn, templateHeight, templateWidth);
        TextBox clonedBox = worksheet.TextBoxes[newIndex];

        // Copy visual properties from the template (optional, can be extended)
        clonedBox.Font.Name = templateBox.Font.Name;
        clonedBox.Font.Size = templateBox.Font.Size;
        clonedBox.Font.IsBold = templateBox.Font.IsBold;
        clonedBox.Font.Color = templateBox.Font.Color;
        clonedBox.Fill.SolidFill.Color = templateBox.Fill.SolidFill.Color;
        clonedBox.Line.Weight = templateBox.Line.Weight;
        clonedBox.Line.DashStyle = templateBox.Line.DashStyle;

        // Update the placeholder text with actual content
        clonedBox.Text = "John Doe";

        // -------------------------------------------------
        // 3. Save the workbook
        // -------------------------------------------------
        workbook.Save("ClonedTextBoxDemo.xlsx");
    }
}
