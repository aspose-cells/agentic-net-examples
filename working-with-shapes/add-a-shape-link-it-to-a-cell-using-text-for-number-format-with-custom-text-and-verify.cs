// Title: Add a label shape linked to a custom‑formatted cell with Aspose.Cells for .NET
// Description: Creates a workbook, writes a numeric value to A1, applies a custom number format (e.g., 0.00 "USD"), adds a label shape at B2, links the shape to the formatted cell, updates the shape to show the formatted text, prints verification details, and saves the file as ShapeLinkedCell.xlsx.
// Keywords: Aspose.Cells label shape | link shape to cell | custom number format | display formatted value in shape | LinkedCell property | C# Aspose.Cells example | Excel shape verification
// Common Searches: Aspose.Cells link textbox to cell with custom format | How to display formatted cell value in a shape using Aspose.Cells | Update linked shape after applying number format Aspose.Cells | C# Aspose.Cells add label shape linked to cell
// Developer Intent: Add a label shape, bind it to a cell that uses a custom number format, refresh the shape to reflect the formatted value, and confirm the link works.
// Use Cases: Financial reports where a shape shows a formatted total amount. | Excel dashboards that use shapes to display live, formatted metrics. | Automated workbook generation with linked shapes for printable summaries.
// AI Prompts: Generate C# code with Aspose.Cells that adds a rectangle shape linked to cell B5 and formats the cell as "dd-MMM-yyyy". | Show how to link multiple shapes to different cells, each with its own custom number format, and verify the displayed values. | Explain how to programmatically confirm that a shape's displayed text matches the cell's formatted string in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, writes a numeric value to A1, applies a custom number format (e.g., 0.00 "USD"), adds a label shape at B2, links the shape to the formatted cell, updates the shape to show the formatted text, prints verification details, and saves the file as ShapeLinkedCell.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value into cell A1
            worksheet.Cells["A1"].PutValue(1234.56);

            // Apply a custom number format (e.g., 0.00 "USD")
            Style customStyle = workbook.CreateStyle();
            customStyle.Custom = "0.00\" USD\"";
            worksheet.Cells["A1"].SetStyle(customStyle);

            // Add a label (text box) shape positioned at cell B2
            // Parameters: upperLeftRow, upperLeftColumn, top, left, height, width
            // Height and width are in pixels; adjust as needed.
            Label shape = worksheet.Shapes.AddLabel(1, 1, 0, 0, 30, 150);

            // Link the shape to the formatted cell A1
            shape.LinkedCell = "$A$1";

            // Refresh the shape so it displays the linked cell's value
            shape.UpdateSelectedValue();

            // Verification output
            Console.WriteLine("Shape's LinkedCell: " + shape.LinkedCell);
            Console.WriteLine("Cell A1 formatted text: " + worksheet.Cells["A1"].StringValue);

            // Save the workbook
            string outputPath = "ShapeLinkedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to: " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
