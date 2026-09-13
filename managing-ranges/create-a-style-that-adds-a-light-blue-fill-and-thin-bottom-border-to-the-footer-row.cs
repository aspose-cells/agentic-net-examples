// Title: Add a light‑blue fill and thin bottom border to a footer row with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a style with a solid light‑blue background and a thin black bottom border, then apply it to a designated footer row. | Provide a complete C# example that defines the footer style, assigns it to cells in row 5 (e.g., columns A‑D), and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# set solid light blue fill for a specific row | How to add a thin bottom border to a footer row in Aspose.Cells workbook | C# Aspose.Cells apply custom style to row range and save file | Create footer row style with background color and border using Aspose.Cells .NET | Set cell style for footer row columns A to D Aspose.Cells example
// Tags: create style with background color Aspose.Cells | apply thin bottom border Aspose.Cells | footer row formatting C# Aspose.Cells | set solid fill for cell range Aspose.Cells .NET | save workbook with styled footer Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The program creates a new workbook, defines a style with a solid light‑blue fill and a thin black bottom border, applies this style to cells in the footer row (row 5, columns A‑D), and saves the file as FooterStyle.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Define the style for the footer row
        Style footerStyle = workbook.CreateStyle();

        // Light blue fill
        footerStyle.ForegroundColor = Color.LightBlue;
        footerStyle.Pattern = BackgroundType.Solid;

        // Thin bottom border
        footerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
        footerStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

        // Apply the style to the footer row (example: row index 5, columns A-D)
        int footerRowIndex = 5;
        for (int col = 0; col <= 3; col++)
        {
            Cell cell = sheet.Cells[footerRowIndex, col];
            cell.SetStyle(footerStyle);
        }

        // Save the workbook
        workbook.Save("FooterStyle.xlsx");
    }
}
