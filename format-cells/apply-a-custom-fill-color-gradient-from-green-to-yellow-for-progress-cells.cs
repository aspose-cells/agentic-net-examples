// Title: Apply a vertical green‑to‑yellow two‑color gradient to Excel cells with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, inserts task names and progress percentages, defines a vertical two‑color gradient (green → yellow) using Aspose.Cells, applies it to cells B2:B4, and saves the file as ProgressCellGradient.xlsx.
// Keywords: Aspose.Cells C# gradient fill | two‑color gradient Excel .NET | vertical gradient cell style | green to yellow background Aspose | set cell style Aspose.Cells
// Common Searches: Aspose.Cells vertical gradient fill C# | apply green yellow gradient to Excel cells .NET | two‑color gradient style Aspose.Cells example | how to set gradient background for a range in Aspose.Cells
// Developer Intent: Generate a workbook and style a specific range with a vertical green‑to‑yellow two‑color gradient using Aspose.Cells for .NET.
// Use Cases: Visualize task completion levels with a gradient that shifts from green (low) to yellow (high). | Create status reports where progress percentages are highlighted by a smooth color transition. | Design lightweight dashboards that emphasize varying completion rates through cell background gradients.
// AI Prompts: Show how to change the gradient direction to horizontal while keeping the green‑to‑yellow colors in Aspose.Cells (C#). | Provide code to apply a three‑color gradient (red, yellow, green) based on numeric values using Aspose.Cells for .NET. | Explain how to assign different gradient variants to multiple cell ranges in the same worksheet.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, inserts task names and progress percentages, defines a vertical two‑color gradient (green → yellow) using Aspose.Cells, applies it to cells B2:B4, and saves the file as ProgressCellGradient.xlsx.
class ProgressCellGradientDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample progress values (0-100)
        sheet.Cells["A1"].PutValue("Task");
        sheet.Cells["B1"].PutValue("Progress");
        sheet.Cells["A2"].PutValue("Task 1");
        sheet.Cells["A3"].PutValue("Task 2");
        sheet.Cells["A4"].PutValue("Task 3");
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["B3"].PutValue(55);
        sheet.Cells["B4"].PutValue(90);

        // Create a style with a two‑color gradient (green → yellow)
        Style gradientStyle = workbook.CreateStyle();
        gradientStyle.SetTwoColorGradient(
            Color.Green,          // start color
            Color.Yellow,         // end color
            GradientStyleType.Vertical, // gradient direction
            1);                   // variant (1‑4)

        // Apply the gradient style to the progress cells (B2:B4)
        for (int row = 2; row <= 4; row++)
        {
            Cell cell = sheet.Cells[row, 1]; // column B (index 1)
            cell.SetStyle(gradientStyle);
        }

        // Save the workbook
        workbook.Save("ProgressCellGradient.xlsx");
    }
}
