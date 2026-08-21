// Title: C# – Apply a Two‑Color Gradient Fill to an Excel Cell with Aspose.Cells
// Description: Creates a workbook, defines a style with a horizontal two‑color gradient (LightBlue → DarkBlue), applies it to cell A1, adjusts row height and column width, and saves the file as CellGradientDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells gradient fill C# | two color cell background .NET | horizontal gradient Excel Aspose | SetTwoColorGradient example | Excel cell style gradient C# | Aspose.Cells formatting tutorial
// Common Searches: how to set gradient background for a cell in Aspose.Cells C# | Aspose.Cells SetTwoColorGradient horizontal example | apply two‑color gradient to Excel cell using .NET | C# code for gradient fill in Aspose.Cells workbook
// Developer Intent: Generate a cell style with a horizontal two‑color gradient and apply it to a specific worksheet cell.
// Use Cases: Emphasize header rows in automated reports with a blue gradient. | Create visually distinct dashboard sections by applying gradient backgrounds to key cells. | Represent intensity or progress levels in status cells using gradient shading.
// AI Prompts: Generate C# code to apply a vertical three‑color gradient to a range of cells with Aspose.Cells. | Show how to set a diagonal two‑color gradient on merged cells while keeping existing formatting intact. | Explain how to vary the gradient colors based on a cell's numeric value in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for GradientStyleType

// Creates a workbook, defines a style with a horizontal two‑color gradient (LightBlue → DarkBlue), applies it to cell A1, adjusts row height and column width, and saves the file as CellGradientDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style and set a two‑color gradient fill
            Style style = workbook.CreateStyle();
            // Gradient from LightBlue to DarkBlue, horizontal direction, variant 1
            style.SetTwoColorGradient(Color.LightBlue, Color.DarkBlue, GradientStyleType.Horizontal, 1);
            // Enable gradient rendering for the style
            style.IsGradient = true;

            // Apply the style to a cell
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Gradient Cell");
            cell.SetStyle(style);

            // Adjust row height and column width for better visibility
            worksheet.Cells.SetRowHeight(0, 40);
            worksheet.Cells.SetColumnWidth(0, 25);

            // Save the workbook (lifecycle save)
            workbook.Save("CellGradientDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
