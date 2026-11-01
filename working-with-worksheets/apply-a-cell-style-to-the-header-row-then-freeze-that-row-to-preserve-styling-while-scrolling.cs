// Title: Apply a light‑gray bold header style to the first row and freeze it using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, defines a solid light‑gray background style with bold centered text, applies it to cells A1‑E1, and freezes the first row using the FreezePanes method in Aspose.Cells. | Generate a .NET example that sets a custom header style on the initial worksheet row and locks that row in place while scrolling by calling Worksheet.FreezePanes.
// Common Searches: Aspose.Cells C# apply background color to header row and keep the row visible while scrolling | How to style the first row and freeze the top row in Excel using Aspose.Cells for .NET | C# example for setting bold centered header style and locking the header row with Aspose.Cells
// Tags: header style application Aspose.Cells | worksheet freeze panes Aspose.Cells | solid background color style Aspose.Cells | bold centered header font Aspose.Cells | xlsx file generation Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new workbook, defines a solid light‑gray background style with bold centered text, applies it to cells A1‑E1 as a header, freezes the top row using FreezePanes, and saves the file as HeaderStyledAndFrozen.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a style for the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.Black;
            headerStyle.HorizontalAlignment = TextAlignmentType.Center;

            // Apply the style to the first row (header)
            // Assuming the header spans columns A to E (0 to 4)
            for (int col = 0; col <= 4; col++)
            {
                Cell cell = sheet.Cells[0, col];
                cell.PutValue($"Header{col + 1}");
                cell.SetStyle(headerStyle);
            }

            // Freeze the header row so it stays visible while scrolling
            // Freeze rows above row 1 (i.e., row 0) and no columns
            // Parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook
            workbook.Save("HeaderStyledAndFrozen.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
