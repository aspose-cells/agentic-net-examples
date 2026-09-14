// Title: How to apply a light‑gray solid fill to the header row of an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a Style with a light‑tone solid background and applies it to all cells in the first row of a worksheet using Aspose.Cells. | Generate a method that determines the last used column in a worksheet, sets a predefined header style for row 0 cells, and saves the workbook as an .xlsx file. | Show how to define a cell style, apply it to a range, and persist the workbook with the styled header using Aspose.Cells in .NET.
// Common Searches: aspocells set header row background to light gray c# | c# aspocells apply solid fill to first row of excel sheet | how to find max data column and style header in aspocells workbook | aspocells export styled worksheet to xlsx file using .NET | c# aspocells define cell style and apply to range
// Tags: header row background styling Aspose.Cells | gray cell fill Aspose.Cells | solid pattern style .NET | column range detection Aspose.Cells | export workbook to xlsx Aspose.Cells

using Aspose.Cells;
using System.Drawing;

// The program creates a new workbook, defines a style with a light‑gray solid background, determines the used column range, applies the style to every cell in the first row, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define a style for the header row
        Style headerStyle = workbook.CreateStyle();
        // Set background color to light gray
        headerStyle.ForegroundColor = Color.LightGray;
        headerStyle.Pattern = BackgroundType.Solid;

        // Determine the last used column in the sheet
        int lastColumn = sheet.Cells.MaxDataColumn;
        // If the sheet is empty, define a default range (e.g., first 5 columns)
        if (lastColumn < 0) lastColumn = 4;

        // Apply the style to each cell in the header row (row index 0)
        for (int col = 0; col <= lastColumn; col++)
        {
            Cell cell = sheet.Cells[0, col];
            cell.SetStyle(headerStyle);
        }

        // Save the workbook (lifecycle rule)
        workbook.Save("Output.xlsx");
    }
}
