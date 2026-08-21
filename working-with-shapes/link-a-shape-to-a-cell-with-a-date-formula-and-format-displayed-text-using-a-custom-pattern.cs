// Title: Aspose.Cells for .NET: Link a label shape to a TODAY() cell and apply a custom date format
// Description: Demonstrates how to create a workbook, insert a TODAY() formula in cell B2, format the cell with the pattern dd‑mmm‑yyyy, add a label shape, bind the shape to the formatted cell using SetLinkedCell, and save the result as LinkedShapeDate.xlsx.
// Keywords: Aspose.Cells | C# | .NET | label shape | linked cell | SetLinkedCell | TODAY formula | custom date format | Excel shape binding | date pattern dd-mmm-yyyy | Excel automation
// Common Searches: link shape to cell Aspose.Cells C# | apply custom date format to cell Aspose.Cells | SetLinkedCell example .NET | add label shape with date formula Aspose.Cells | bind Excel shape to TODAY() cell using Aspose
// Developer Intent: Create a worksheet where a label shape displays the current date by linking it to a cell that contains a TODAY() formula and uses a custom date format.
// Use Cases: Dynamic report headers that automatically show the current date. | Dashboard templates where shapes reflect formatted date values. | Invoice or receipt templates with a linked date shape that updates on each open.
// AI Prompts: Generate C# code that links a label shape to a cell with a TODAY() formula and formats the cell as dd‑mmm‑yyyy using Aspose.Cells. | Explain how SetLinkedCell works for binding shapes to cells in Aspose.Cells for .NET. | Show an example of adding a label shape, applying a custom date format, and linking it to a date cell in an Excel workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, insert a TODAY() formula in cell B2, format the cell with the pattern dd‑mmm‑yyyy, add a label shape, bind the shape to the formatted cell using SetLinkedCell, and save the result as LinkedShapeDate.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Place a date formula in cell B2
            worksheet.Cells["B2"].Formula = "=TODAY()";

            // Apply a custom date format to the cell (e.g., 25-Dec-2023)
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-mmm-yyyy";
            worksheet.Cells["B2"].SetStyle(dateStyle);

            // Add a label shape to the worksheet
            // Parameters: upper left row, upper left column, lower right row, lower right column, height, width
            // Height and width are in points; adjust as needed.
            Label label = worksheet.Shapes.AddLabel(4, 1, 6, 3, 30, 100);

            // Link the label shape to the cell containing the date formula
            // Formula is in A1‑style, not R1C1, and uses the local language settings
            label.SetLinkedCell("$B$2", false, true);

            // Save the workbook
            workbook.Save("LinkedShapeDate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
