// Title: Link a Label Shape to a Percentage‑Formatted Cell with Aspose.Cells for .NET
// Description: Creates a new workbook, writes 0.25 to cell B2, applies the built‑in percent format (index 10) which sets IsPercent = true, adds a label shape at row 2 column 2 (100 × 50 pt), links the shape to B2 using SetLinkedCell, refreshes the displayed value, prints the label text, and saves the file as ShapeLinkedPercent.xlsx.
// Keywords: Aspose.Cells label shape | link shape to cell | percentage number format | IsPercent property | SetLinkedCell C# | Excel shape binding | Aspose.Cells .NET example
// Common Searches: Aspose.Cells link label to cell percentage | How to bind a shape to a cell in C# | Check IsPercent after applying number format Aspose.Cells | Add label shape and display cell value Aspose.Cells | Set linked cell for a shape Aspose.Cells .NET
// Developer Intent: Add a label shape, bind it to a cell formatted as a percent, and verify the displayed text.
// Use Cases: Dynamic dashboards where shapes reflect live percentage calculations. | Automated report generation with shapes that automatically show formatted values. | Testing that a cell’s number format is recognized as percent before linking to a shape.
// AI Prompts: Generate C# code that adds a rectangle shape linked to cell C5 formatted as currency and confirms the format. | Show how to link multiple shapes to different cells, each using a custom number format, and output their texts. | Explain the parameters of SetLinkedCell, how they affect linking behavior, and how to refresh shape content after the source cell changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, writes 0.25 to cell B2, applies the built‑in percent format (index 10) which sets IsPercent = true, adds a label shape at row 2 column 2 (100 × 50 pt), links the shape to B2 using SetLinkedCell, refreshes the displayed value, prints the label text, and saves the file as ShapeLinkedPercent.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a numeric value (0.25) into cell B2 (represents 25%)
            Cell cell = worksheet.Cells["B2"];
            cell.PutValue(0.25);

            // Apply a built‑in percentage number format (index 10) which sets IsPercent = true
            Style style = cell.GetStyle();
            style.Number = 10; // Built‑in percent format
            cell.SetStyle(style);

            // Verify that the style reports IsPercent = true
            Console.WriteLine("IsPercent after applying format: " + style.IsPercent);

            // Add a label shape at row 2, column 2 with size 100x50 points
            // AddLabel overload requires six integer parameters in this API version:
            // upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, height, width
            Shape shape = worksheet.Shapes.AddLabel(2, 2, 2, 2, 100, 50);
            Label label = (Label)shape;

            // Link the label to cell B2 so it displays the cell's value
            label.SetLinkedCell("$B$2", false, true);
            label.UpdateSelectedValue(); // Refresh the displayed value

            // Retrieve and display the text shown by the shape
            string displayedText = label.Text; // Use the Text property for label content
            Console.WriteLine("Label displays: " + displayedText);

            // Save the workbook
            workbook.Save("ShapeLinkedPercent.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
