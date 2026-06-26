using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsCustomStyleExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("Input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a custom style to highlight customer names
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.ForegroundColor = Color.Yellow;          // cell background
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.Font.IsBold = true;                     // bold font
            highlightStyle.Font.Color = Color.Red;                 // font color
            highlightStyle.Font.Size = 12;

            // Assume that cells mapped to /Customer/Name are in column B (index 1)
            // Apply the style to rows 2 through 10 (adjust as needed)
            for (int row = 1; row <= 9; row++) // zero‑based index: row 1 = second row
            {
                Cell nameCell = cells[row, 1]; // column B
                // Optionally verify the cell contains a name before styling
                if (!string.IsNullOrEmpty(nameCell.StringValue))
                {
                    nameCell.SetStyle(highlightStyle);
                }
            }

            // Save the modified workbook
            workbook.Save("Output.xlsx");
        }
    }
}