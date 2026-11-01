// Title: How to add a thick left border and light‑blue background to a single cell using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a style with a black thick left border and a light‑blue fill, applies it to cell A1, and saves the workbook as an .xlsx file using Aspose.Cells. | Write a C# method that receives a worksheet and a cell address, then sets a thick left border and a light‑blue background on that cell with Aspose.Cells styling APIs.
// Common Searches: Aspose.Cells C# set left border thickness for a specific cell | C# Aspose.Cells apply light blue background color to a cell | How to style a single cell with border and fill using Aspose.Cells .NET | Example of creating a custom cell style in an Aspose.Cells workbook C#
// Tags: Aspose.Cells set cell left border thickness C# | Aspose.Cells background color fill C# | custom cell style Aspose.Cells workbook | StyleFlag apply all attributes Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// // Creates a workbook, defines a style with a thick left border and light‑blue background, applies it to cell A1, and saves the file as StyledCell.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the cell to which the style will be applied (e.g., A1)
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Styled Cell");

        // Create a new style object
        Style style = workbook.CreateStyle();

        // Set a thick left border
        style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
        style.Borders[BorderType.LeftBorder].Color = Color.Black; // optional border color

        // Set a light blue background fill
        style.ForegroundColor = Color.LightBlue;
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the cell
        StyleFlag flag = new StyleFlag();
        flag.All = true; // apply all style attributes
        cell.SetStyle(style, flag);

        // Save the workbook to a file
        workbook.Save("StyledCell.xlsx");
    }
}
