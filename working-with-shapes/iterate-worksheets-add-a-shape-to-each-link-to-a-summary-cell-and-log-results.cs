// Title: Add a Rectangle Shape to Every Worksheet and Link It to B1 Using Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds sample sheets with a summary value in B1, iterates through all worksheets, inserts a rectangle shape on each sheet, sets the shape's name and text, links the shape to the sheet's B1 cell, logs the shape name and linked value, and saves the file as ShapesLinked.xlsx.
// Keywords: Aspose.Cells C# shape example | add rectangle to worksheet Aspose.Cells | link shape to cell B1 | iterate worksheets Aspose.Cells | shape logging Aspose.Cells | save workbook with linked shapes | .NET spreadsheet shape linking | Aspose.Cells shape automation
// Common Searches: How to add a rectangle shape to each sheet with Aspose.Cells .NET | Aspose.Cells link shape to a specific cell | Iterate worksheets and insert shapes in C# | Log shape creation details in Aspose.Cells | Set linked cell for a shape using Aspose.Cells
// Developer Intent: Insert a rectangle shape on every worksheet, bind it to the sheet's B1 cell, and output a console log with the shape name and linked value.
// Use Cases: Build a dashboard where each sheet has a clickable rectangle that reflects the sheet's summary value. | Automate report templates by programmatically adding linked shapes to key data cells across multiple worksheets. | Debug and verify shape insertion by printing the linked cell value during workbook generation.
// AI Prompts: Generate C# code with Aspose.Cells that adds a circle shape to each worksheet, links it to cell C5, and logs the shape name and linked value. | Update the provided example to set the rectangle's fill color to light blue and add a hyperlink to another sheet while keeping the B1 link. | Create a reusable method that accepts a Workbook, adds a custom button shape to every sheet, links it to a specified summary cell, and returns a list of the created Shape objects.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that creates a workbook, adds sample sheets with a summary value in B1, iterates through all worksheets, inserts a rectangle shape on each sheet, sets the shape's name and text, links the shape to the sheet's B1 cell, logs the shape name and linked value, and saves the file as ShapesLinked.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Prepare sample worksheets with a summary cell (B1)
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.Cells["A1"].PutValue("Summary");
        ws1.Cells["B1"].PutValue(10);

        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
        ws2.Cells["A1"].PutValue("Summary");
        ws2.Cells["B1"].PutValue(20);

        // Iterate through each worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Add a rectangle shape to the worksheet
            // Parameters: upperLeftRow, top, upperLeftColumn, left, height, width
            Shape shape = ws.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);
            shape.Name = $"Rect_{ws.Name}";
            shape.Text = $"Linked to {ws.Name}!B1";

            // Link the shape to the summary cell B1 of the same worksheet
            // The two boolean flags are for row/column linking; false/false means standard cell link
            shape.SetLinkedCell($"{ws.Name}!B1", false, false);

            // Log the operation
            Console.WriteLine($"Added shape '{shape.Name}' to worksheet '{ws.Name}' linked to cell B1 (value = {ws.Cells["B1"].Value})");
        }

        // Save the workbook (lifecycle: save)
        workbook.Save("ShapesLinked.xlsx");
    }
}
